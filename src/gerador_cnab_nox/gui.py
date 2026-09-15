from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import PhotoImage, StringVar, Tk, filedialog, messagebox, ttk

from .advogado_rules import load_rules as load_lawyer_rules
from .errors import UserFacingError, ValidationError
from .failure_aliases import load_aliases
from .gui_aliases import AliasEditor
from .gui_messages import explain_issue, explain_pending, money
from .gui_theme import configure_theme
from .gui_widgets import Card, ResultPanel, ScrollableBody, paragraph
from .models import CESSAO, NORMAL, PfmiInput
from .service import generate_cnab, prepare_workbook
from .workbook import read_intermediate

NORMAL_LABEL = "Normal"
CESSAO_LABEL = "Cessão da Cessão"
MODALITY_BY_LABEL = {NORMAL_LABEL: NORMAL, CESSAO_LABEL: CESSAO}
LABEL_BY_MODALITY = {value: key for key, value in MODALITY_BY_LABEL.items()}
STEPS = ("Configurar lote", "Preparar e revisar Excel", "Gerar TXT")


@dataclass
class _PfmiRow:
    path: str
    modalidade: str = NORMAL
    comissao_texto: str = ""
    cessao_ja_configurada: bool = False
    emissao_nova_cessao_texto: str = ""


@dataclass
class _PreparedReview:
    result: object
    rows: tuple
    error: Exception | None = None


def _prepare_review(*args, **kwargs):
    """Read the saved review data in the worker; this function never uses Tk."""
    result = prepare_workbook(*args, **kwargs)
    try:
        rows = tuple(read_intermediate(result.output_path).rows)
    except Exception as exc:
        # The workbook was saved: a summary failure must not claim otherwise.
        return _PreparedReview(result, (), exc)
    return _PreparedReview(result, rows)


def _parse_preenchimentos_automaticos(raw: object) -> list[dict]:
    """PREENCHIMENTOS_AUTOMATICOS is a JSON audit trail written at prepare
    time (matching.py); malformed/empty content never breaks the summary."""
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except (TypeError, ValueError):
        return []
    return [item for item in parsed if isinstance(item, dict)] if isinstance(parsed, list) else []


def _composition_message(c: dict) -> str:
    total_pfmi = money(c["total_pfmi"]) if c["total_pfmi"] is not None else "indisponível"
    total_analitico = (
        money(c["total_analitico"]) if c["total_analitico"] is not None else "indisponível"
    )
    diferenca = money(c["diferenca"]) if c["diferenca"] is not None else "indisponível"
    nominal = (
        money(c["nominal_analitico"]) if c["nominal_analitico"] is not None else "indisponível"
    )
    if c["estado"] == "PROPOSTA":
        prefixo = "Composição encontrada — confira antes de aprovar"
        rodape = (
            "Esta sugestão não confirma identidade: confira o instrumento e aprove "
            "explicitamente em COMPOSICAO_APROVADA no Excel; enquanto isso os títulos "
            "permanecem pendentes."
        )
    elif c["estado"] == "BLOQUEADA_CREDITO_NAO_LOCALIZADO":
        prefixo = "Composição bloqueada"
        rodape = (
            "Não localizamos o crédito automaticamente. Veja os candidatos em "
            "COMPOSICAO_CANDIDATOS (nome, documento mascarado, valor e linha) e, se "
            "você já conferiu o instrumento, informe a aba e a linha exatas do "
            "Analítico em COMPOSICAO_SELECAO_MANUAL_ABA/COMPOSICAO_SELECAO_MANUAL_LINHA "
            "(iguais em todas as linhas da composição) no Excel; só é aceito se fechar "
            "exatamente com a soma dos componentes."
        )
    else:
        prefixo = "Composição bloqueada"
        rodape = (
            "Esta sugestão não confirma identidade: confira o instrumento e aprove "
            "explicitamente em COMPOSICAO_APROVADA no Excel; enquanto isso os títulos "
            "permanecem pendentes."
        )
    return (
        f"{prefixo}: {c['participantes']} · total PFMI {total_pfmi} · "
        f"total Analítico {total_analitico} · diferença {diferenca} · "
        f"nominal do crédito {nominal} · motivo: {c['motivo']}. {rodape}"
    )


class Application:
    def __init__(self, root: Tk):
        self.root = root
        self.fonts = configure_theme(root)
        self.root.title("Gerador CNAB NOX")
        screen_w, screen_h = root.winfo_screenwidth(), root.winfo_screenheight()
        self.root.geometry(f"{min(1040, screen_w - 60)}x{min(700, screen_h - 100)}")
        self.root.minsize(min(900, screen_w - 60), min(600, screen_h - 100))

        self.analytic = StringVar(master=root)
        self.due = StringVar(master=root)
        self.liquidation = StringVar(master=root)
        self.sequence = StringVar(master=root)
        self.intermediate = StringVar(master=root)
        self.prepared_path = StringVar(master=root)
        self.generated_path = StringVar(master=root)
        self.pfmi_count = StringVar(master=root, value="0 PFMIs")
        self.pfmi_modality = StringVar(master=root, value=NORMAL_LABEL)
        self.pfmi_commission = StringVar(master=root)
        self.pfmi_emissao_cessao = StringVar(master=root)
        self.selected_path = StringVar(master=root)
        self.selected_name = StringVar(master=root, value="Selecione uma PFMI para configurar.")
        self.configuration_summary = StringVar(master=root)
        self.configuration_notice = StringVar(master=root)
        self.activity = StringVar(master=root, value="Etapa 1 de 3 · Configure o lote.")
        self.status = StringVar(master=root)
        self.is_busy = False
        self._alias_revision = 0
        self.current_step = 1
        self._pfmi_rows: dict[str, _PfmiRow] = {}
        self._next_pfmi_id = 1
        self._editing_pfmi_id: str | None = None
        self._loading_pfmi_editor = False
        self._prepared_signature = None
        self._job_queue: queue.Queue = queue.Queue()
        self._poll_id = None
        self._worker = None
        self._opener_jobs: list[tuple[subprocess.Popen, int, str]] = []
        self._opener_poll_id = None
        self._focus_reveal_id = None
        self._controls: list[tuple[ttk.Widget, str]] = []
        self._build()
        self.pfmi_commission.trace_add("write", self._on_commission_edited)
        self.pfmi_emissao_cessao.trace_add("write", self._on_emissao_cessao_edited)
        for variable in (self.analytic, self.due, self.liquidation, self.sequence):
            variable.trace_add("write", self._configuration_changed)
        self.intermediate.trace_add("write", self._intermediate_changed)
        self.root.protocol("WM_DELETE_WINDOW", self._request_close)
        self.root.bind("<FocusIn>", self._reveal_focus, add="+")
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.root.bind(event, self.body.wheel, add="+")
        self._configuration_changed()
        self._set_pfmi_editor_enabled(False)
        self.show_step(1)
        self._update_actions()

    def _button(self, parent, text, command, *, primary=False):
        widget = ttk.Button(
            parent,
            text=text,
            command=command,
            style="Primary.TButton" if primary else "TButton",
        )
        self._controls.append((widget, "normal"))
        return widget

    def _entry(self, parent, variable, *, width=None, readonly=False):
        kwargs = {
            "textvariable": variable,
            "state": "readonly" if readonly else "normal",
            "font": self.fonts["body"],
        }
        if width:
            kwargs["width"] = width
        widget = ttk.Entry(parent, **kwargs)
        if not readonly:
            self._controls.append((widget, "normal"))
        return widget

    def _build(self):
        self.container = ttk.Frame(self.root, style="Shell.TFrame")
        self.container.pack(fill="both", expand=True)
        header = ttk.Frame(self.container, style="Header.TFrame", padding=(28, 14))
        header.pack(fill="x")
        brand = ttk.Frame(header, style="Header.TFrame")
        brand.pack(side="left")
        self.title_label = ttk.Label(brand, text="Gerador CNAB NOX", style="Header.TLabel")
        self.title_label.pack(anchor="w")
        ttk.Label(
            brand, text="Do lote ao arquivo pronto para envio", style="HeaderHelp.TLabel"
        ).pack(anchor="w", pady=(4, 0))
        self.direct_excel_button = self._button(
            header, "Já tenho um Excel revisado", lambda: self.show_step(3)
        )
        self.direct_excel_button.pack(side="right")
        self.direct_excel_button.configure(style="Header.TButton")
        self.new_batch_button = self._button(header, "Novo Lote", self._new_batch)
        self.new_batch_button.pack(side="right", padx=(0, 8))
        self.new_batch_button.configure(style="Header.TButton")
        navigation = ttk.Frame(self.container, style="Shell.TFrame", padding=(24, 12))
        navigation.pack(fill="x")
        self.step_buttons = []
        icon_size = 45 if float(self.root.tk.call("tk", "scaling")) > 1.7 else 30
        self.step_icons = {
            (index, state): PhotoImage(
                master=self.root,
                file=Path(__file__).with_name("assets") / f"step_{index}_{state}_{icon_size}.png",
            )
            for index in range(1, 4)
            for state in ("active", "idle")
        }
        for index, title in enumerate(STEPS, 1):
            title = "Preparar e revisar" if index == 2 else title
            button = self._button(navigation, " " + title, lambda i=index: self.show_step(i))
            button.configure(image=self.step_icons[index, "idle"], compound="left")
            column = (index - 1) * 2
            button.grid(row=0, column=column, sticky="ew")
            navigation.columnconfigure(column, weight=1)
            if index < 3:
                ttk.Separator(navigation, orient="horizontal").grid(
                    row=0, column=column + 1, sticky="ew", padx=8
                )
                navigation.columnconfigure(column + 1, minsize=28)
            self.step_buttons.append(button)

        self.body = ScrollableBody(self.container)
        self.body.pack(fill="both", expand=True, padx=24)
        self.pages = {}
        for step in range(1, 4):
            page = ttk.Frame(self.body.content, style="Shell.TFrame")
            page.grid(row=0, column=0, sticky="nsew")
            self.pages[step] = page
        self._build_configuration()
        self._build_review()
        self._build_generation()

        footer = ttk.Frame(self.container, style="Shell.TFrame", padding=(24, 12))
        # Reserve the actions before the Canvas consumes the remaining height at high DPI.
        footer.pack(side="bottom", fill="x", before=self.body)
        self.activity_label = ttk.Label(
            footer, textvariable=self.activity, style="Shell.TLabel", wraplength=700
        )
        self.progress = ttk.Progressbar(footer, mode="indeterminate")
        self.footers = {}
        for step in range(1, 4):
            actions = ttk.Frame(footer, style="Shell.TFrame")
            actions.pack(fill="x")
            self.footers[step] = actions
        self.continue_button = self._button(
            self.footers[1], "Continuar", lambda: self.show_step(2), primary=True
        )
        self.continue_button.pack(side="right")
        self.back_config_button = self._button(self.footers[2], "Voltar", lambda: self.show_step(1))
        self.back_config_button.pack(side="left")
        self.back_config_button.configure(style="Footer.TButton")
        self.prepare_button = self._button(
            self.footers[2], "Preparar Excel", self._prepare, primary=True
        )
        self.prepare_button.pack(side="right")
        self.continue_generate_button = self._button(
            self.footers[2], "Continuar para gerar TXT", lambda: self.show_step(3)
        )
        self.open_excel_button = self._button(
            self.footers[2],
            "Abrir Excel para revisar",
            lambda: self._open_local(self.prepared_path.get(), 2),
            primary=True,
        )
        self.back_review_button = self._button(self.footers[3], "Voltar", lambda: self.show_step(2))
        self.back_review_button.pack(side="left")
        self.back_review_button.configure(style="Footer.TButton")
        self.generate_button = self._button(
            self.footers[3], "Validar e gerar TXT", self._generate, primary=True
        )
        self.generate_button.pack(side="right")
        self.open_generated_folder_button = self._button(
            self.footers[3],
            "Abrir pasta",
            lambda: self._open_local(self.generated_path.get(), 3, folder=True),
            primary=True,
        )
        self.new_batch_after_generate_button = self._button(
            self.footers[3], "Novo Lote", self._new_batch
        )
        self.results = {2: self.review_result, 3: self.generation_result}
        self.review_result.show("Pronto para preparar", "Escolha onde salvar o Excel ao preparar.")
        self._reset_generation()

    def _build_configuration(self):
        page = self.pages[1]
        paragraph(page, "Configure o lote", style="PageTitle.TLabel", pady=(0, 6))
        self.instructions_label = paragraph(
            page,
            "Adicione as PFMIs na ordem desejada e confira a modalidade de cada arquivo.",
            style="PageHelp.TLabel",
            pady=(0, 18),
        )
        file_card = Card(page)
        file_card.pack(fill="x", pady=(0, 16))
        bar = ttk.Frame(file_card)
        bar.pack(fill="x", pady=(4, 8))
        ttk.Label(bar, text="PFMIs do lote", style="Section.TLabel").pack(side="left")
        self.pfmi_count_label = ttk.Label(bar, textvariable=self.pfmi_count)
        self.pfmi_count_label.pack(side="left", padx=12)
        self.add_pfmi_button = self._button(bar, "Selecionar PFMIs", self._add_pfmis)
        self.add_pfmi_button.pack(side="right")
        self.add_more_pfmi_button = self._button(
            bar, "Adicionar mais PFMIs", lambda: self._add_pfmis(replace=False)
        )
        self.add_more_pfmi_button.pack(side="right", padx=(0, 8))
        table = ttk.Frame(file_card)
        table.pack(fill="x")
        table.columnconfigure(0, weight=1)
        self.pfmi_tree = ttk.Treeview(
            table,
            columns=("ordem", "arquivo", "modalidade", "comissao"),
            show="headings",
            height=3,
            selectmode="browse",
        )
        for key, label in (
            ("ordem", "Ordem"),
            ("arquivo", "Arquivo PFMI"),
            ("modalidade", "Modalidade"),
            ("comissao", "Comissão %"),
        ):
            self.pfmi_tree.heading(key, text=label, anchor="w")
        measure = self.fonts["body"].measure
        self.pfmi_tree.column("ordem", width=measure("Ordem") + 24, stretch=False, anchor="center")
        self.pfmi_tree.column("arquivo", width=360, minwidth=170, stretch=True)
        self.pfmi_tree.column(
            "modalidade", width=measure(CESSAO_LABEL) + 28, stretch=False, minwidth=150
        )
        self.pfmi_tree.column(
            "comissao", width=measure("Comissão %") + 24, stretch=False, anchor="center"
        )
        vertical = ttk.Scrollbar(table, command=self.pfmi_tree.yview)
        self.pfmi_tree.configure(yscrollcommand=vertical.set)
        self.pfmi_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        self.pfmi_tree.bind("<<TreeviewSelect>>", self._on_pfmi_selected)
        self.pfmi_tree.bind("<Button-1>", lambda _e: "break" if self.is_busy else None)
        self.pfmi_tree.bind("<Key>", lambda _e: "break" if self.is_busy else None)
        tools = ttk.Frame(file_card)
        tools.pack(fill="x", pady=(8, 12))
        for attribute, title, command in (
            ("remove_pfmi_button", "Remover", self._remove_pfmi),
            ("move_up_button", "Mover acima", lambda: self._move_pfmi(-1)),
            ("move_down_button", "Mover abaixo", lambda: self._move_pfmi(1)),
        ):
            button = self._button(tools, title, command)
            button.configure(style="Quiet.TButton")
            button.pack(side="left", padx=(0, 8))
            setattr(self, attribute, button)
        editor_card = Card(page)
        editor_card.pack(fill="x", pady=(0, 16))
        paragraph(editor_card, variable=self.selected_name, style="Strong.TLabel", pady=(0, 6))
        self.selected_path_entry = self._entry(editor_card, self.selected_path, readonly=True)
        self.selected_path_entry.pack(fill="x", pady=(0, 10))
        settings = ttk.Frame(editor_card)
        settings.pack(fill="x")
        ttk.Label(settings, text="Modalidade").grid(row=0, column=0, sticky="w")
        ttk.Label(settings, text="Comissão (%)").grid(row=0, column=1, sticky="w", padx=(16, 0))
        ttk.Label(settings, text="Data de emissão da nova cessão (dd/mm/aaaa)").grid(
            row=0, column=2, sticky="w", padx=(16, 0)
        )
        self.modality_combo = ttk.Combobox(
            settings,
            textvariable=self.pfmi_modality,
            values=(NORMAL_LABEL, CESSAO_LABEL),
            state="readonly",
            width=24,
            font=self.fonts["body"],
        )
        self.modality_combo.grid(row=1, column=0, sticky="ew", pady=(4, 6))
        self.modality_combo.bind("<<ComboboxSelected>>", self._on_modality_changed)
        self._controls.append((self.modality_combo, "readonly"))
        self.commission_entry = self._entry(settings, self.pfmi_commission, width=14)
        self.commission_entry.grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(4, 6))
        self.emissao_cessao_entry = self._entry(settings, self.pfmi_emissao_cessao, width=14)
        self.emissao_cessao_entry.grid(row=1, column=2, sticky="w", padx=(16, 0), pady=(4, 6))
        self.commission_help_label = paragraph(
            editor_card,
            "Estas opções se aplicam somente ao arquivo selecionado. "
            "Em Cessão, a taxa inicial é 15%; zero e campo apagado são preservados. "
            "A data de emissão da nova cessão vale para todos os títulos deste arquivo "
            "e nunca é preenchida automaticamente.",
        )
        data_card = Card(page)
        data_card.pack(fill="x")
        paragraph(data_card, "Dados do lote", style="Section.TLabel")
        self.aliases_button = self._button(
            data_card, "Equivalências de falência", self._edit_aliases
        )
        self.aliases_button.pack(anchor="w", pady=(4, 8))
        sources = ttk.Frame(data_card)
        sources.pack(fill="x")
        sources.columnconfigure(1, weight=1)
        self.analytic_entry, self.analytic_select_button = self._file_row(
            sources,
            0,
            "Prospect Analítico",
            self.analytic,
            [("Analítico", "*.xls *.xlsx *.csv")],
        )
        self.due_entry, self.due_select_button = self._file_row(
            sources,
            1,
            "Base de Vencimentos",
            self.due,
            [("Excel", "*.xlsx")],
        )
        params = ttk.Frame(data_card)
        params.pack(fill="x", pady=(10, 0))
        ttk.Label(params, text="Data de liquidação (dd/mm/aaaa)").grid(row=0, column=0, sticky="w")
        ttk.Label(params, text="Primeira sequência").grid(row=0, column=1, sticky="w", padx=(20, 0))
        self.liquidation_entry = self._entry(params, self.liquidation, width=23)
        self.liquidation_entry.grid(row=1, column=0, sticky="w", pady=(4, 8))
        self.sequence_entry = self._entry(params, self.sequence, width=18)
        self.sequence_entry.grid(row=1, column=1, sticky="w", padx=(20, 0), pady=(4, 8))
        paragraph(
            data_card,
            "Confira a sequência no estoque do fundo. Data e sequência podem ser "
            "corrigidas no RESUMO; comissão inválida exige uma nova preparação.",
        )

    def _build_review(self):
        page = self.pages[2]
        paragraph(page, "Prepare e revise o Excel", style="PageTitle.TLabel", pady=(0, 6))
        paragraph(
            page,
            "Prepare o arquivo e faça a conferência no seu editor de planilhas.",
            style="PageHelp.TLabel",
            pady=(0, 18),
        )
        self.summary_label = paragraph(
            page, variable=self.configuration_summary, style="PageHelp.TLabel"
        )
        self.notice_label = paragraph(
            page, variable=self.configuration_notice, style="warning.TLabel"
        )
        self.review_result = ResultPanel(page, self.fonts["body"])
        self.review_result.pack(fill="x", pady=(2, 8))
        self.prepared_files = ttk.Frame(self.review_result)
        paragraph(self.prepared_files, "Último Excel salvo", style="Strong.TLabel", pady=(0, 4))
        self.prepared_path_entry = self._entry(
            self.prepared_files, self.prepared_path, readonly=True
        )
        self.prepared_path_entry.pack(fill="x", pady=(0, 8))
        buttons = ttk.Frame(self.prepared_files)
        buttons.pack(fill="x", pady=(0, 12))
        self.open_prepared_folder_button = self._button(
            buttons,
            "Abrir pasta",
            lambda: self._open_local(self.prepared_path.get(), 2, folder=True),
        )
        self.open_prepared_folder_button.pack(side="left")
        self.prepare_again_button = self._button(buttons, "Preparar novo Excel", self._prepare)
        self.prepare_again_button.pack(side="left", padx=(8, 0))
        self.review_guide = Card(page)
        self.review_guide.pack(fill="x", pady=(16, 0))
        paragraph(self.review_guide, "O que conferir no Excel", style="Section.TLabel")
        for line in (
            "1. RESUMO: confira a data de liquidação e a primeira sequência.",
            "2. CREDITOS: confira INCLUIR_CNAB e os campos finais; corrija as pendências.",
            "3. PENDENCIAS_PFMI: confira o que exige corrigir a origem e preparar novamente.",
            "4. Salve o Excel antes de continuar para gerar o TXT.",
        ):
            paragraph(self.review_guide, line, style="TLabel", pady=(0, 6))
        paragraph(
            self.review_guide,
            "Amarelo: campos editáveis. Azul: referências protegidas. "
            "APROVADO=SIM confirma somente divergência de nome, respeitando a modalidade.",
            pady=(8, 0),
        )
        self.opening_review_notice = paragraph(page, style="warning.TLabel", pady=(10, 0))

    def _build_generation(self):
        page = self.pages[3]
        paragraph(page, "Gere o TXT", style="PageTitle.TLabel", pady=(0, 6))
        paragraph(
            page,
            "Selecione o Excel revisado e salvo. A geração usa somente esse arquivo.",
            style="PageHelp.TLabel",
            pady=(0, 18),
        )
        selector = Card(page)
        selector.pack(fill="x", pady=(0, 12))
        selector.columnconfigure(1, weight=1)
        self.intermediate_entry, self.intermediate_select_button = self._file_row(
            selector, 0, "Excel revisado", self.intermediate, [("Excel", "*.xlsx")]
        )
        paragraph(
            page,
            "Se você salvou uma cópia com outro nome, selecione essa cópia aqui.",
            style="PageHelp.TLabel",
        )
        self.generation_result = ResultPanel(page, self.fonts["body"])
        self.generation_result.pack(fill="x", pady=(6, 12))
        self.generated_files = ttk.Frame(self.generation_result)
        paragraph(self.generated_files, "TXT salvo", style="Strong.TLabel", pady=(0, 4))
        self.generated_path_entry = self._entry(
            self.generated_files, self.generated_path, readonly=True
        )
        self.generated_path_entry.pack(fill="x", pady=(0, 8))
        self.opening_generation_notice = paragraph(page, style="warning.TLabel", pady=(10, 0))

    def _file_row(self, parent, row, label, variable, filetypes):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
        entry = self._entry(parent, variable)
        entry.grid(row=row, column=1, sticky="ew", pady=5)
        button = self._button(parent, "Selecionar...", lambda: self._select(variable, filetypes))
        button.grid(row=row, column=2, padx=(8, 0), pady=5)
        return entry, button

    def show_step(self, step):
        if self.is_busy:
            return
        previous_focus = self.root.focus_get()
        self.current_step = step
        for index, page in self.pages.items():
            page.grid() if index == step else page.grid_remove()
        for index, footer in self.footers.items():
            footer.pack(fill="x") if index == step else footer.pack_forget()
        for index, button in enumerate(self.step_buttons, 1):
            button.configure(style="ActiveStep.TButton" if index == step else "Step.TButton")
            button.configure(image=self.step_icons[index, "active" if index == step else "idle"])
        self.activity.set(f"Etapa {step} de 3 · {STEPS[step - 1]}.")
        self.body.canvas.yview_moveto(0)
        self._sync_status()
        # Navigation moves keyboard focus into the visible stage.
        self.root.update_idletasks()
        if previous_focus is None:
            self.title_label.focus_set()
        elif self.body.contains(previous_focus) or not previous_focus.winfo_ismapped():
            self.step_buttons[step - 1].focus_set()

    def _reveal_focus(self, _event=None):
        if self._focus_reveal_id:
            self.root.after_cancel(self._focus_reveal_id)
        self._focus_reveal_id = self.root.after_idle(self._reveal_current_focus)

    def _reveal_current_focus(self):
        self._focus_reveal_id = None
        current = self.root.focus_get()
        if current is not None:
            self.body.reveal(current)

    def _sync_status(self):
        panel = self.results.get(self.current_step)
        self.status.set(
            f"{panel.title.cget('text')}\n{panel.body.cget('text')}\n{panel.metric_summary}".strip()
            if panel
            else "Adicione as PFMIs para começar."
        )

    def _signature(self):
        return (
            tuple(
                (p.path, p.modalidade, p.comissao_texto, p.emissao_nova_cessao_texto)
                for p in self._ordered_pfmi_inputs()
            ),
            self.analytic.get(),
            self.due.get(),
            self.liquidation.get(),
            self.sequence.get(),
            self._alias_revision,
        )

    def _configuration_changed(self, *_args):
        inputs = self._ordered_pfmi_inputs()
        analytic_name = Path(self.analytic.get()).name or "não selecionado"
        self.configuration_summary.set(
            f"PFMIs: {len(inputs)} · Analítico: {analytic_name}\n"
            f"Base: {Path(self.due.get()).name or 'não selecionada'}\n"
            f"Liquidação: {self.liquidation.get() or 'a informar no RESUMO'} · "
            f"Primeira sequência: {self.sequence.get() or 'a informar no RESUMO'}"
        )
        changed = (
            self._prepared_signature is not None and self._signature() != self._prepared_signature
        )
        self.configuration_notice.set(
            "Para aplicar estas mudanças, prepare um novo Excel. O arquivo anterior permanece "
            "independente da configuração atual."
            if changed
            else ""
        )
        if changed:
            self.notice_label.pack(fill="x", before=self.review_result, pady=(0, 10))
        else:
            self.notice_label.pack_forget()

    def _reset_generation(self):
        self.generated_path.set("")
        self.generated_files.pack_forget()
        self.open_generated_folder_button.pack_forget()
        self.new_batch_after_generate_button.pack_forget()
        self.generate_button.configure(text="Validar e gerar TXT", style="Primary.TButton")
        self.generate_button.pack(side="right")
        self.generation_result.show(
            "Excel aguardando validação",
            "Revise e salve os campos finais antes de clicar em Validar e gerar TXT.",
        )
        self.opening_generation_notice.configure(text="")

    def _reset_batch_state(self):
        """Start a brand-new batch: clear every PFMI, source, date/sequence
        and prepared/generated result, so residual state from a previous
        test (e.g. a different liquidation date) never mixes into the next
        one. Never called implicitly by ordinary back/forward navigation -
        only by an explicit "Novo Lote" action - so reviewing or tweaking
        the current batch's configuration is never destructive."""
        if self.is_busy:
            return
        for item_id in list(self.pfmi_tree.get_children()):
            self.pfmi_tree.delete(item_id)
        self._pfmi_rows.clear()
        self._next_pfmi_id = 1
        self._clear_pfmi_editor()
        self._refresh_pfmi_tree()
        self.analytic.set("")
        self.due.set("")
        self.liquidation.set("")
        self.sequence.set("")
        self._prepared_signature = None
        self.prepared_path.set("")
        self.intermediate.set("")
        self.prepared_files.pack_forget()
        self.prepare_button.pack(side="right")
        self.open_excel_button.pack_forget()
        self.continue_generate_button.pack_forget()
        self.summary_label.pack(fill="x", anchor="w", pady=(0, 10), before=self.review_result)
        self.review_result.show(
            "Pronto para preparar", "Escolha onde salvar o Excel ao preparar."
        )
        self._reset_generation()
        self._drain_job_queue()
        self._configuration_changed()
        self._sync_status()

    def _drain_job_queue(self):
        # Only reachable when not busy (see the is_busy guard above), so the
        # worker never puts anything new while this runs; defensive only.
        while True:
            try:
                self._job_queue.get_nowait()
            except queue.Empty:
                break

    def _new_batch(self):
        self._reset_batch_state()
        self.show_step(1)

    def _intermediate_changed(self, *_args):
        self._reset_generation()
        self._sync_status()

    def _select(self, variable, filetypes):
        if self.is_busy:
            return
        chosen = filedialog.askopenfilename(
            parent=self.root, filetypes=[*filetypes, ("Todos", "*.*")]
        )
        if chosen:
            variable.set(chosen)

    def _add_pfmis(self, *, replace=True):
        if self.is_busy:
            return
        chosen = filedialog.askopenfilenames(
            parent=self.root,
            title="Selecionar PFMIs" if replace else "Adicionar mais PFMIs",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if not chosen:
            return
        if replace and self._pfmi_rows:
            # Default action starts a fresh selection instead of silently
            # accumulating files from a previous, possibly unrelated batch
            # (e.g. mixing PFMIs from two different liquidation dates).
            # "Adicionar mais PFMIs" is the explicit, deliberate way to
            # combine several files into the same batch (e.g. Normal +
            # Cessão da Cessão for the same test).
            for item_id in list(self.pfmi_tree.get_children()):
                self.pfmi_tree.delete(item_id)
            self._pfmi_rows.clear()
            self._editing_pfmi_id = None
        last_id = None
        for path in chosen:
            item_id = f"PFMI_{self._next_pfmi_id:04d}"
            self._next_pfmi_id += 1
            self._pfmi_rows[item_id] = _PfmiRow(path=str(path))
            self.pfmi_tree.insert("", "end", iid=item_id)
            last_id = item_id
        if last_id is not None:
            self._refresh_pfmi_tree()
            self.pfmi_tree.selection_set(last_id)
            self.pfmi_tree.focus(last_id)
            self.pfmi_tree.see(last_id)
            self._load_pfmi_editor(last_id)

    def _remove_pfmi(self):
        if self.is_busy:
            return
        selected = self.pfmi_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        old_index = list(self.pfmi_tree.get_children()).index(item_id)
        self.pfmi_tree.delete(item_id)
        self._pfmi_rows.pop(item_id)
        self._editing_pfmi_id = None
        self._refresh_pfmi_tree()
        remaining = self.pfmi_tree.get_children()
        if remaining:
            next_id = remaining[min(old_index, len(remaining) - 1)]
            self.pfmi_tree.selection_set(next_id)
            self.pfmi_tree.focus(next_id)
            self._load_pfmi_editor(next_id)
        else:
            self._clear_pfmi_editor()

    def _move_pfmi(self, offset):
        if self.is_busy:
            return
        selected = self.pfmi_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        children = list(self.pfmi_tree.get_children())
        destination = children.index(item_id) + offset
        if 0 <= destination < len(children):
            self.pfmi_tree.move(item_id, "", destination)
            self.pfmi_tree.see(item_id)
            self._refresh_pfmi_tree()
            self._update_actions()

    def _on_pfmi_selected(self, _event=None):
        if self.is_busy:
            return
        selected = self.pfmi_tree.selection()
        if selected:
            self._load_pfmi_editor(selected[0])
        else:
            self._clear_pfmi_editor()

    def _load_pfmi_editor(self, item_id):
        row = self._pfmi_rows[item_id]
        self._loading_pfmi_editor = True
        try:
            self._editing_pfmi_id = item_id
            self.pfmi_modality.set(LABEL_BY_MODALITY[row.modalidade])
            self.pfmi_commission.set(row.comissao_texto)
            self.pfmi_emissao_cessao.set(row.emissao_nova_cessao_texto)
            self.selected_name.set(f"Editando: {Path(row.path).name}")
            self.selected_path.set(row.path)
        finally:
            self._loading_pfmi_editor = False
        self._set_pfmi_editor_enabled(True)
        self._update_actions()

    def _clear_pfmi_editor(self):
        self._loading_pfmi_editor = True
        try:
            self._editing_pfmi_id = None
            self.pfmi_modality.set(NORMAL_LABEL)
            self.pfmi_commission.set("")
            self.pfmi_emissao_cessao.set("")
            self.selected_name.set("Selecione uma PFMI para configurar.")
            self.selected_path.set("")
        finally:
            self._loading_pfmi_editor = False
        self._set_pfmi_editor_enabled(False)
        self._update_actions()

    def _set_pfmi_editor_enabled(self, enabled):
        enabled = enabled and not self.is_busy
        self.modality_combo.configure(state="readonly" if enabled else "disabled")
        row = self._pfmi_rows.get(self._editing_pfmi_id or "")
        self.commission_entry.configure(
            state="normal" if enabled and row and row.modalidade == CESSAO else "disabled"
        )
        self.emissao_cessao_entry.configure(
            state="normal" if enabled and row and row.modalidade == CESSAO else "disabled"
        )

    def _on_modality_changed(self, _event=None):
        if self.is_busy or self._loading_pfmi_editor or not self._editing_pfmi_id:
            return
        row = self._pfmi_rows[self._editing_pfmi_id]
        row.modalidade = MODALITY_BY_LABEL[self.pfmi_modality.get()]
        if row.modalidade == CESSAO and not row.cessao_ja_configurada:
            row.cessao_ja_configurada = True
            self.pfmi_commission.set("15")
        self._set_pfmi_editor_enabled(True)
        self._refresh_pfmi_tree()

    def _on_commission_edited(self, *_args):
        if self.is_busy or self._loading_pfmi_editor or not self._editing_pfmi_id:
            return
        self._pfmi_rows[self._editing_pfmi_id].comissao_texto = self.pfmi_commission.get()
        self._refresh_pfmi_tree()

    def _on_emissao_cessao_edited(self, *_args):
        if self.is_busy or self._loading_pfmi_editor or not self._editing_pfmi_id:
            return
        self._pfmi_rows[self._editing_pfmi_id].emissao_nova_cessao_texto = (
            self.pfmi_emissao_cessao.get()
        )
        self._refresh_pfmi_tree()

    def _refresh_pfmi_tree(self):
        children = self.pfmi_tree.get_children()
        for order, item_id in enumerate(children, 1):
            row = self._pfmi_rows[item_id]
            self.pfmi_tree.item(
                item_id,
                values=(
                    order,
                    Path(row.path).name,
                    LABEL_BY_MODALITY[row.modalidade],
                    row.comissao_texto if row.modalidade == CESSAO else "—",
                ),
            )
        count = len(children)
        self.pfmi_count.set(f"{count} PFMI" if count == 1 else f"{count} PFMIs")
        self._configuration_changed()

    def _ordered_pfmi_inputs(self):
        return [
            PfmiInput(
                self._pfmi_rows[item_id].path,
                self._pfmi_rows[item_id].modalidade,
                self._pfmi_rows[item_id].comissao_texto,
                self._pfmi_rows[item_id].emissao_nova_cessao_texto,
            )
            for item_id in self.pfmi_tree.get_children()
        ]

    def _update_actions(self):
        selected = self.pfmi_tree.selection()
        children = self.pfmi_tree.get_children()
        index = children.index(selected[0]) if selected else -1
        for button, possible in (
            (self.remove_pfmi_button, index >= 0),
            (self.move_up_button, index > 0),
            (self.move_down_button, 0 <= index < len(children) - 1),
        ):
            button.configure(state="normal" if possible and not self.is_busy else "disabled")

    def _edit_aliases(self):
        if self.is_busy:
            return
        try:
            self.alias_editor = AliasEditor(self.root, self._aliases_saved)
        except UserFacingError as exc:
            messagebox.showerror("Equivalências indisponíveis", str(exc), parent=self.root)

    def _aliases_saved(self):
        self._alias_revision += 1
        self._configuration_changed()

    def _prepare(self):
        if self.is_busy:
            return
        self.show_step(2)
        pfmis = self._ordered_pfmi_inputs()
        if not pfmis or not self.analytic.get().strip() or not self.due.get().strip():
            self.review_result.show(
                "Arquivos obrigatórios",
                "Adicione ao menos uma PFMI e selecione o Analítico e a Base de Vencimentos.",
                kind="warning",
            )
            self._sync_status()
            return
        destination = filedialog.asksaveasfilename(
            parent=self.root,
            title="Salvar Excel para revisão",
            defaultextension=".xlsx",
            initialfile="excel_cnab_nox.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not destination:
            return
        # Capture every Tk value on its owning thread before dispatch.
        try:
            aliases = load_aliases()
            lawyer_rules = load_lawyer_rules()
        except UserFacingError as exc:
            self.review_result.show("Equivalências indisponíveis", str(exc), kind="error")
            self._sync_status()
            return
        args = (pfmis, self.analytic.get(), self.due.get())
        kwargs = {
            "liquidation_date": self.liquidation.get(),
            "first_sequence": self.sequence.get(),
            "output_path": destination,
            "failure_aliases": aliases,
            "lawyer_rules": lawyer_rules,
        }
        signature = self._signature()
        self._start_job(
            2,
            "Preparando o Excel. Aguarde a conclusão.",
            lambda: _prepare_review(*args, **kwargs),
            lambda result: self._prepared(result, len(pfmis), signature),
        )

    def _prepared(self, result, pfmi_count, signature):
        review = result if isinstance(result, _PreparedReview) else None
        if review is not None:
            result = review.result
        self._prepared_signature = signature
        self.prepared_path.set(str(result.output_path))
        self.intermediate.set(str(result.output_path))
        self._configuration_changed()
        groups: dict[str, dict] = {}
        preenchimentos_comprovados: list[str] = []
        sugestoes_aguardando: list[str] = []
        if review is not None:
            if review.error is not None:
                raise review.error
            loaded_rows = review.rows
        else:
            # Direct calls used by layout previews; real jobs already carry the rows.
            try:
                loaded_rows = read_intermediate(result.output_path).rows
            except UserFacingError:
                loaded_rows = ()
        for row in loaded_rows:
            v = row.values
            g = groups.setdefault(
                v["ID_GRUPO"], {"credor": v["NOME_CEDENTE_PFMI"], "ok": False, "pendencias": ""}
            )
            if (
                v["STATUS"] == "OK" and v.get("INCLUIR_CNAB") == "SIM"
                and (not v.get("COMPOSICAO_ID") or v.get("COMPOSICAO_APROVADA") == "SIM")
            ):
                g["ok"] = True
            elif not g["ok"]:
                # A blank PENDENCIAS cell round-trips through openpyxl as
                # None, not "" - reachable now that a composition's only
                # pendency can be cleared by the nominal suggestion while the
                # row still isn't approved/included.
                g["pendencias"] = v["PENDENCIAS"] or ""
            preenchimentos_linha = _parse_preenchimentos_automaticos(
                v.get("PREENCHIMENTOS_AUTOMATICOS")
            )
            for preenchimento in preenchimentos_linha:
                campo = preenchimento.get("campo", "Campo não informado")
                linha = (
                    f"{v['NOME_CEDENTE_PFMI']}: {campo} = "
                    f"{preenchimento.get('valor_sugerido', '')} "
                    f"({preenchimento.get('regra', 'Regra não informada')}"
                )
                source = preenchimento.get("fonte")
                linha += f", fonte: {source})" if source else ")"
                if preenchimento.get("exige_aprovacao_humana"):
                    sugestoes_aguardando.append(linha)
                else:
                    preenchimentos_comprovados.append(linha)
        total = len(groups)
        prontas = sum(1 for g in groups.values() if g["ok"])
        pendentes = [g for g in groups.values() if not g["ok"]]
        body = f"{total} operações · {prontas} prontas · {len(pendentes)} pendentes."
        pending_lines = tuple(
            f"{g['credor']}: {explain_pending(g['pendencias'])}" for g in pendentes
        )
        suggestion_lines = tuple(
            "Sugestão (confira o instrumento, não confirma identidade): "
            f"{s['credor_a']} ({money(s['valor_a'])}) + {s['credor_b']} ({money(s['valor_b'])}) "
            f"= {money(s['total'])} · possível crédito no Analítico: "
            + ", ".join(c["referencia"] for c in s["creditos_encontrados"])
            for s in result.related_suggestions
        )
        composition_lines = tuple(
            _composition_message(c) for c in result.compositions
        )
        aguardando = sum(1 for c in result.compositions if c["estado"] == "PROPOSTA")
        titulos_envolvidos = sum(len(c["component_ids"]) for c in result.compositions)
        # Três blocos separados visualmente (regra "Preenchimentos comprovados"
        # vs "Decisões que exigem conferência" vs pendências sem solução
        # automática) dentro do mesmo painel de texto já existente - nenhuma
        # etapa nova, só a reorganização do resumo pós-preparo.
        sections = ["── Preenchimentos comprovados (automáticos) ──"]
        sections.extend(preenchimentos_comprovados or ["Nenhum preenchimento nesta categoria."])
        sugestoes = tuple(sugestoes_aguardando) + suggestion_lines + composition_lines
        sections.append("── Sugestões aguardando confirmação ──")
        sections.extend(sugestoes or ["Nenhuma sugestão aguardando confirmação."])
        sections.append("── Pendências sem solução automática ──")
        sections.extend(pending_lines or ["Nenhuma pendência nesta categoria."])
        self.review_result.show(
            "Excel preparado",
            body,
            kind="success" if not pendentes else "warning",
            issues=tuple(sections),
            details=result.warnings,
            metrics=(
                ("operações", total),
                ("prontas", prontas),
                ("pendentes", len(pendentes)),
                ("composições encontradas", len(result.compositions)),
                ("títulos envolvidos", titulos_envolvidos),
                ("composições aguardando aprovação", aguardando),
            ),
        )
        # Six cards in two rows stay readable at the supported minimum width.
        for column in range(6):
            self.review_result.metrics_frame.columnconfigure(
                column, weight=1 if column < 3 else 0, uniform="metrics" if column < 3 else ""
            )
        for index, block in enumerate(self.review_result.metrics_frame.winfo_children()):
            block.grid_configure(row=index // 3, column=index % 3, pady=(0, 8))
        self._show_prepared_files()
        self.root.update_idletasks()

    def _show_prepared_files(self):
        self.summary_label.pack_forget()
        anchor = self.review_result.metrics_frame
        if not anchor.winfo_manager():
            anchor = self.review_result.body
        self.prepared_files.pack(fill="x", after=anchor)
        self.prepare_button.pack_forget()
        self.open_excel_button.pack(side="right")
        self.continue_generate_button.pack(side="right", padx=(0, 8))
        self.opening_review_notice.configure(text="")
        self.body.canvas.yview_moveto(0)

    def _generate(self):
        if self.is_busy:
            return
        self.show_step(3)
        if not self.intermediate.get().strip():
            self.generation_result.show(
                "Excel obrigatório", "Selecione o Excel revisado e salvo.", kind="warning"
            )
            self._sync_status()
            return
        destination = filedialog.asksaveasfilename(
            parent=self.root,
            title="Salvar TXT CNAB",
            defaultextension=".txt",
            initialfile="cnab_nox.txt",
            filetypes=[("CNAB TXT", "*.txt")],
        )
        if not destination:
            return
        source = self.intermediate.get()
        analytic_path = self.analytic.get().strip() or None
        self._reset_generation()
        self._start_job(
            3,
            "Validando o Excel e gerando o TXT. Aguarde.",
            lambda: generate_cnab(
                source,
                output_path=destination,
                lawyer_rules=load_lawyer_rules(),
                analytic_path=analytic_path,
            ),
            self._generated,
        )

    def _generated(self, result):
        self.generated_path.set(str(result.output_path))
        composition_lines = tuple(
            "Composição gerada — confira antes de enviar: "
            + "; ".join(f"{c['nome']} = R$ {money(c['vl_nominal'])}" for c in comp["componentes"])
            + f" · total nominal do grupo: R$ {money(comp['total_nominal'])}"
            for comp in result.compositions
        )
        self.generation_result.show(
            "TXT salvo com sucesso",
            f"Sequência: {result.first_sequence}–{result.last_sequence} · "
            f"Avisos: {result.warning_count}",
            kind="success",
            issues=tuple(explain_issue(v) for v in result.warnings) + composition_lines,
            details=result.warnings,
            metrics=(
                ("Títulos", result.detail_count),
                ("Nominal", f"R$ {money(result.total_nominal)}"),
                ("Aquisição", f"R$ {money(result.total_present)}"),
            ),
        )
        self.generated_files.pack(
            fill="x", after=self.generation_result.metrics_frame, pady=(0, 12)
        )
        self.generate_button.pack_forget()
        self.open_generated_folder_button.pack(side="right")
        self.new_batch_after_generate_button.pack(side="right", padx=(0, 8))
        self.generate_button.configure(text="Gerar outro TXT", style="TButton")
        self.generate_button.pack(side="right", padx=(0, 8))
        self.body.canvas.yview_moveto(0)

    def _start_job(self, step, text, operation, completed):
        if self.is_busy:
            return
        self._busy(True, text)
        self.results[step].show(text, kind="info")
        self._sync_status()

        def work():
            try:
                value = operation()
                self._job_queue.put((True, value))
            except Exception as exc:  # noqa: BLE001 - worker transports errors to the UI
                self._job_queue.put((False, exc))

        self._job_completed = completed
        self._job_step = step
        self._worker = threading.Thread(target=work, name="cnab-nox-operation", daemon=False)
        try:
            self._worker.start()
        except Exception as exc:  # noqa: BLE001 - recover controls if thread cannot start
            self._show_error(exc, step)
            self._busy(False)
            return
        self._poll_id = self.root.after(40, self._poll_job)

    def _poll_job(self):
        self._poll_id = None
        try:
            success, value = self._job_queue.get_nowait()
        except queue.Empty:
            self._poll_id = self.root.after(40, self._poll_job)
            return
        # The queue is the worker's only contact with the UI. Even after() is
        # scheduled on the Tk thread, so no worker relies on Tcl cross-thread calls.
        self._poll_id = self.root.after(0, self._finish_job, success, value)

    def _finish_job(self, success, value):
        self._poll_id = None
        try:
            if success:
                self._job_completed(value)
            else:
                self._show_error(value, self._job_step)
        except Exception as exc:
            if success and self._job_step == 2:
                result = value.result if isinstance(value, _PreparedReview) else value
                self.prepared_path.set(str(result.output_path))
                self.intermediate.set(str(result.output_path))
                self.review_result.show(
                    "Excel salvo; resumo indisponível",
                    "O arquivo foi salvo, mas não foi possível apresentar o resumo. "
                    "Abra o Excel para revisar as informações.",
                    kind="warning",
                    details=(f"Falha ao apresentar resumo: {type(exc).__name__}",),
                )
                self._show_prepared_files()
            else:
                self._show_error(exc, self._job_step)
        finally:
            self._busy(False)
            self._sync_status()
            self.body.reveal(self.results[self._job_step].title, align_top=True)
            self.root.update_idletasks()

    def _busy(self, busy, text=None):
        self.is_busy = busy
        for widget, idle_state in self._controls:
            widget.configure(state="disabled" if busy else idle_state)
        self.pfmi_tree.state(["disabled"] if busy else ["!disabled"])
        if busy:
            self.activity_label.pack(fill="x", before=self.footers[self.current_step], pady=(0, 6))
            self.progress.pack(fill="x", before=self.footers[self.current_step], pady=(0, 8))
            self.progress.start(15)
            self.activity.set(text or "Aguarde a conclusão.")
        else:
            self.activity_label.pack_forget()
            self.progress.stop()
            self.progress.pack_forget()
            self.activity.set(f"Etapa {self.current_step} de 3 · {STEPS[self.current_step - 1]}.")
        self._set_pfmi_editor_enabled(self._editing_pfmi_id is not None)
        self._update_actions()

    def _show_error(self, exc, step):
        panel = self.results[step]
        if isinstance(exc, ValidationError):
            panel.show(
                "O TXT não foi gerado.",
                "Corrija os apontamentos abaixo, salve o Excel e tente novamente.",
                kind="error",
                issues=tuple(explain_issue(v) for v in exc.issues),
                details=tuple(exc.issues),
            )
        elif isinstance(exc, UserFacingError):
            panel.show(
                "O TXT não foi gerado." if step == 3 else "Não foi possível preparar o Excel.",
                str(exc),
                kind="error",
            )
        else:
            panel.show(
                "Não foi possível concluir",
                "Ocorreu um erro inesperado. Os arquivos originais não foram alterados. "
                "Confira o destino e se algum arquivo de saída está aberto. "
                "Verifique a pasta antes de tentar novamente.",
                kind="error",
            )
        self.body.canvas.yview_moveto(0)

    def _request_close(self):
        if self.is_busy:
            self.activity.set("Aguarde a conclusão da operação para fechar a janela.")
            return
        if self._opener_poll_id:
            self.root.after_cancel(self._opener_poll_id)
        if self._focus_reveal_id:
            self.root.after_cancel(self._focus_reveal_id)
        self.root.destroy()

    def _open_local(self, value, step, *, folder=False):
        if self.is_busy or not value:
            return
        path = Path(value).absolute()
        path = path.parent if folder else path
        label = self.opening_review_notice if step == 2 else self.opening_generation_notice
        label.configure(text="")
        if not path.exists():
            label.configure(
                text="O arquivo ou a pasta não está mais neste caminho. Confira o destino."
            )
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            else:
                command = "open" if sys.platform == "darwin" else "xdg-open"
                process = subprocess.Popen(
                    [command, str(path)],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                self._opener_jobs.append((process, step, str(path)))
                if not self._opener_poll_id:
                    self._opener_poll_id = self.root.after(100, self._poll_openers)
        except OSError:
            self._open_failed(step)

    def _open_failed(self, step):
        label = self.opening_review_notice if step == 2 else self.opening_generation_notice
        label.configure(
            text="Não foi possível abrir pelo sistema. Copie o caminho acima e abra manualmente "
            "no seu editor de planilhas ou gerenciador de arquivos. O arquivo salvo foi preservado."
        )

    def _poll_openers(self):
        self._opener_poll_id = None
        pending = []
        for process, step, path in self._opener_jobs:
            result = process.poll()
            if result is None:
                pending.append((process, step, path))
            elif result != 0:
                self._open_failed(step)
        self._opener_jobs = pending
        if pending:
            self._opener_poll_id = self.root.after(100, self._poll_openers)


def main():
    root = Tk()
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
