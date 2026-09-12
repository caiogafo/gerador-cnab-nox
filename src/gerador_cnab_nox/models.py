from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

FINAL_FIELDS = (
    "DOC_CEDENTE",
    "NOME_CEDENTE",
    "SEU_NUMERO",
    "NU_DOCUMENTO",
    "DT_VENCIMENTO",
    "VL_NOMINAL",
    "DOC_SACADO",
    "NOME_SACADO",
    "VL_PRESENTE",
    "TIPO_PESSOA_SACADO",
    "ENDERECO",
    "CEP",
    "TP_TITULO",
    "DT_EMISSAO_TITULO",
    "COOBRIGACAO",
    "TIPO_PESSOA_CEDENTE",
    "NFE",
    "VALOR_PAGO_TITULO",
    "INDEXADOR",
    "TAXA_INDEXADOR",
    "MOVIMENTO",
)

CONTROL_FIELDS = (
    "ID_LINHA",
    "ID_BLOCO",
    "ID_GRUPO",
    "REFERENCIA_ANALITICO",
    "INCLUIR_CNAB",
    "APROVADO",
    "STATUS",
    "PENDENCIAS",
    "ALERTAS",
    "TOTAL_PFMI_GRUPO",
    "NOME_CEDENTE_PFMI",
    "FALENCIA_PFMI",
    "DOC_BENEFICIARIO_PFMI",
    "PENDENCIAS_FONTE_PFMI",
    "ID_CREDITO",
    "ID_ARQUIVO",
    "ORDEM_ARQUIVO",
    "ARQUIVO_PFMI",
    "HASH_PFMI",
    "ABA_PFMI",
    "LINHAS_PFMI",
    "ORDEM_GRUPO",
    "PAGAMENTOS_GRUPO",
    "MODALIDADE",
    "COMISSAO_TEXTO",
    "PERCENTUAL_USADO",
    "PENDENCIAS_PARAMETROS",
    "FALENCIA_CANONICA",
    "HASH_ANALITICO",
    "ABA_ANALITICO",
    "LINHA_ANALITICO",
    "NOME_CEDENTE_ANALITICO",
    "DOC_CEDENTE_ANALITICO",
    "MOTIVO_CORRESPONDENCIA",
    "AQUISICAO_ORIGINAL",
    "NOMINAL_ORIGINAL",
    "ASSINATURA_ORIGINAL",
    "BASE_COMISSAO",
    "COMISSAO_CALCULADA",
    "AQUISICAO_ESPERADA",
    "ORIGEM_DOC_CEDENTE",
    "ESTADO_CNPJ_CONEXCRED",
    "ORIGENS_CNPJ_CONEXCRED",
    "MOTIVO_SUBSTITUICAO",
    "NOME_ESPERADO",
    "EVIDENCIA_BASE",
    "ORIGINAL_ANALITICO",
    "DIFERENCA_PRESENTE",
    "DIFERENCA_NOMINAL",
    "NOME_CEDENTE_TECNICO",
    "NOME_SACADO_TECNICO",
    "ACAO_NECESSARIA",
    "EXIGE_NOVA_PREPARACAO",
    "EMISSAO_NOVA_CESSAO_TEXTO_ORIGINAL",
)

EDITABLE_CONTROLS = {"INCLUIR_CNAB", "APROVADO"}
DISPLAY_CONTROLS = {
    "STATUS",
    "PENDENCIAS",
    "ALERTAS",
    "DIFERENCA_PRESENTE",
    "DIFERENCA_NOMINAL",
    "NOME_CEDENTE_TECNICO",
    "NOME_SACADO_TECNICO",
    "ACAO_NECESSARIA",
    "EXIGE_NOVA_PREPARACAO",
}
IMMUTABLE_FIELDS = tuple(
    key for key in CONTROL_FIELDS if key not in EDITABLE_CONTROLS | DISPLAY_CONTROLS
)
NORMAL = "NORMAL"
CESSAO = "CESSAO_DA_CESSAO"


@dataclass(frozen=True)
class PfmiInput:
    path: str | Path
    modalidade: str = NORMAL
    comissao_texto: str = "15"
    emissao_nova_cessao_texto: str = ""


@dataclass(frozen=True)
class PreparationRequest:
    pfmis: list[PfmiInput]
    analitico_path: str | Path
    vencimentos_path: str | Path
    data_liquidacao: Any = None
    primeira_sequencia: Any = None


@dataclass(frozen=True)
class PfmiPayment:
    block_id: str
    source_row: int
    title: str
    failure: str
    cedent_name: str
    beneficiary_name: str
    beneficiary_document: str
    operation_value: Decimal | None
    beneficiary_value: Decimal | None
    issues: tuple[str, ...] = ()
    source_sheet: str = ""
    source_file: str = ""
    file_id: str = "F0001"
    file_order: int = 1
    source_hash: str = ""
    modality: str = NORMAL
    commission_text: str = "15"
    emissao_nova_cessao_texto: str = ""
    raw_values: tuple[Any, ...] = ()


@dataclass(frozen=True)
class PfmiUnknownRow:
    block_id: str
    source_row: int
    reason: str
    nonempty_columns: str
    source_sheet: str = ""
    source_file: str = ""
    file_id: str = "F0001"
    raw_values: tuple[Any, ...] = ()


@dataclass
class PfmiData:
    payments: list[PfmiPayment] = field(default_factory=list)
    unknown_rows: list[PfmiUnknownRow] = field(default_factory=list)
    evidence: list[dict[str, Any]] = field(default_factory=list)
    sources: list[dict[str, Any]] = field(default_factory=list)
    operational_digest: str = ""


@dataclass(frozen=True)
class AnalyticCredit:
    source_row: int
    reference: str
    cedent_document: str
    cedent_name: str
    campaign: str
    nominal_value: Decimal | None
    acquisition_value: Decimal | None
    signature_date: date | None
    source_hash: str = ""
    source_sheet: str = ""
    raw_values: dict[str, Any] = field(default_factory=dict)

    @property
    def credit_id(self) -> str:
        import hashlib

        origin = f"{self.source_hash}|{self.source_sheet}|{self.source_row}|{self.reference}"
        return hashlib.sha256(origin.encode()).hexdigest()


@dataclass(frozen=True)
class DueRecord:
    source_row: int
    debtor_name: str
    debtor_document: str
    due_date: date | None
    source_hash: str = ""
    source_sheet: str = ""
    raw_values: dict[str, Any] = field(default_factory=dict)


@dataclass
class PreparedRow:
    values: dict[str, Any]
    originals: dict[str, Any]


@dataclass
class PreparedBatch:
    rows: list[PreparedRow]
    unknown_rows: list[PfmiUnknownRow]
    liquidation_date: date | None
    first_sequence: int | None
    warnings: list[str] = field(default_factory=list)
    payment_count: int = 0
    group_count: int = 0
    evidence: list[dict[str, Any]] = field(default_factory=list)
    sources: list[dict[str, Any]] = field(default_factory=list)
    payments_snapshot: list[dict[str, Any]] = field(default_factory=list)
    related_suggestions: list[dict[str, Any]] = field(default_factory=list)


@dataclass(frozen=True)
class PreparationResult:
    output_path: Path
    row_count: int
    selected_count: int
    warning_count: int
    warnings: tuple[str, ...] = ()
    payment_count: int = 0
    group_count: int = 0
    related_suggestions: tuple[dict[str, Any], ...] = ()


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    detail_count: int
    total_nominal: Decimal
    total_present: Decimal
    first_sequence: int
    last_sequence: int
    warning_count: int
    warnings: tuple[str, ...]
