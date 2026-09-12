"""Local equivalence editor, saved only after explicit operator confirmation."""

from tkinter import BooleanVar, StringVar, Toplevel, ttk

from .errors import InputFileError
from .failure_aliases import load_aliases, save_aliases, validate_aliases


class AliasEditor:
    def __init__(self, parent, on_saved, path=None):
        self.path = path
        self.on_saved = on_saved
        self.aliases = load_aliases(path)
        self.window = Toplevel(parent)
        self.window.title("Equivalências de falência")
        self.window.geometry("760x480")
        self.window.transient(parent)
        self.window.grab_set()
        frame = ttk.Frame(self.window, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame, text="Cadastre somente nomes confirmados da mesma falência."
        ).pack(anchor="w")
        ttk.Label(
            frame, text="Aplicado aos próximos Excel. Arquivos já preparados não mudam."
        ).pack(anchor="w")
        self.tree = ttk.Treeview(frame, columns=("source", "target"), show="headings", height=7)
        self.tree.heading("source", text="Nome alternativo")
        self.tree.heading("target", text="Nome de referência")
        self.tree.pack(fill="both", expand=True, pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.select)
        self.source = StringVar(master=self.window)
        self.target = StringVar(master=self.window)
        fields = (("Nome alternativo", self.source), ("Nome de referência", self.target))
        for label, variable in fields:
            ttk.Label(frame, text=label).pack(anchor="w")
            ttk.Entry(frame, textvariable=variable).pack(fill="x")
        self.confirmed = BooleanVar(master=self.window, value=False)
        ttk.Checkbutton(frame, text="Conferi que os dois nomes representam a mesma falência",
                        variable=self.confirmed).pack(anchor="w", pady=8)
        for variable in (self.source, self.target):
            variable.trace_add("write", lambda *_: self.confirmed.set(False))
        actions = ttk.Frame(frame)
        actions.pack(fill="x")
        ttk.Button(actions, text="Salvar equivalência", command=self.save).pack(side="left")
        ttk.Button(actions, text="Remover selecionada", command=self.remove).pack(
            side="left", padx=8
        )
        ttk.Button(actions, text="Fechar", command=self.window.destroy).pack(side="right")
        self.message = StringVar(master=self.window)
        ttk.Label(frame, textvariable=self.message, wraplength=700).pack(anchor="w", pady=8)
        self.refresh()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for source, target in sorted(self.aliases.items()):
            self.tree.insert("", "end", values=(source, target))

    def select(self, _event=None):
        selected = self.tree.selection()
        if selected:
            source, target = self.tree.item(selected[0], "values")
            self.source.set(source)
            self.target.set(target)

    def persist(self, values):
        try:
            values = validate_aliases(values)
            save_aliases(values, self.path)
        except (InputFileError, OSError) as exc:
            self.message.set(str(exc))
            return
        self.aliases = values
        self.refresh()
        self.confirmed.set(False)
        self.message.set("Cadastro salvo. Prepare um novo Excel para aplicar as equivalências.")
        self.on_saved()

    def save(self):
        if not self.confirmed.get():
            self.message.set("Confirme que os dois nomes representam a mesma falência.")
            return
        try:
            pair = validate_aliases({self.source.get(): self.target.get()})
        except InputFileError as exc:
            self.message.set(str(exc))
            return
        self.persist({**self.aliases, **pair})

    def remove(self):
        selected = self.tree.selection()
        if selected:
            source = self.tree.item(selected[0], "values")[0]
            values = dict(self.aliases)
            del values[source]
            self.persist(values)
