from __future__ import annotations

import hashlib
import json
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from .cnab import write_cnab
from .failure_aliases import validate_aliases
from .matching import prepare_batch
from .models import GenerationResult, PfmiInput, PreparationRequest, PreparationResult
from .normalize import parse_date, parse_positive_int
from .readers import read_analytic, read_due_base, read_pfmis
from .validation import validate_for_generation
from .workbook import read_intermediate, write_intermediate


def prepare_workbook(
    pfmi_path: str | Path | list[PfmiInput] | PreparationRequest,
    analytic_path: str | Path | None = None,
    due_base_path: str | Path | None = None,
    *,
    liquidation_date: Any = None,
    first_sequence: Any = None,
    output_path: str | Path | None = None,
    log_directory: str | Path | None = None,
    failure_aliases: dict[str, str] | None = None,
    lawyer_rules: dict[str, tuple[str, ...]] | None = None,
) -> PreparationResult:
    aliases = validate_aliases(failure_aliases or {})
    if isinstance(pfmi_path, PreparationRequest):
        request = pfmi_path
        inputs = request.pfmis
        analytic_path, due_base_path = request.analitico_path, request.vencimentos_path
        liquidation_date, first_sequence = request.data_liquidacao, request.primeira_sequencia
    else:
        inputs = pfmi_path if isinstance(pfmi_path, list) else [PfmiInput(pfmi_path)]
    if not analytic_path or not due_base_path:
        from .errors import InputFileError

        raise InputFileError("Selecione o Analítico e a Base de Vencimentos")
    sources = [*[Path(item.path) for item in inputs], Path(analytic_path), Path(due_base_path)]
    fingerprints = [_source_fingerprint(path) for path in sources]
    pfmi = read_pfmis(inputs)
    analytic = read_analytic(analytic_path)
    due_base = read_due_base(due_base_path)
    parsed_date, date_warning = _optional_date(liquidation_date)
    parsed_sequence, sequence_warning = _optional_sequence(first_sequence)
    batch = prepare_batch(
        pfmi,
        analytic,
        due_base,
        liquidation_date=parsed_date,
        first_sequence=parsed_sequence,
        failure_aliases=aliases,
        lawyer_rules=lawyer_rules or {},
    )
    batch.liquidation_date = parsed_date if parsed_date is not None else liquidation_date
    batch.first_sequence = parsed_sequence if parsed_sequence is not None else first_sequence
    batch.sources = [
        *pfmi.sources,
        {
            "tipo": "ANALITICO",
            **fingerprints[-2],
            "sha256": hashlib.sha256(Path(analytic_path).read_bytes()).hexdigest(),
        },
        {
            "tipo": "BASE",
            **fingerprints[-1],
            "sha256": hashlib.sha256(Path(due_base_path).read_bytes()).hexdigest(),
        },
        {
            "tipo": "PARAMETROS_LOTE",
            "liquidacao_texto": str(liquidation_date),
            "sequencia_texto": str(first_sequence),
        },
        {"tipo": "EQUIVALENCIAS_FALENCIAS", "confirmadas": aliases},
    ]
    batch.warnings.extend(
        warning for warning in (date_warning, sequence_warning) if warning is not None
    )
    selected_preview = sum(row.values.get("INCLUIR_CNAB") == "SIM" for row in batch.rows)
    if (
        parsed_sequence is not None
        and parsed_sequence + max(0, selected_preview - 1) > 9_999_999_999
    ):
        batch.warnings.append("FAIXA_SEQUENCIA_FORA_LAYOUT")
    destination = (
        Path(output_path) if output_path else _unique_output(Path.cwd(), "excel_cnab_nox", ".xlsx")
    )
    write_intermediate(batch, destination)
    selected = selected_preview
    warnings = (
        sum(row.values.get("STATUS") != "OK" for row in batch.rows)
        + len(batch.unknown_rows)
        + len(batch.warnings)
    )
    operation_warnings: list[str] = []
    try:
        _audit(
            log_directory or destination.parent / "logs",
            {
                "evento": "preparar",
                "entradas": fingerprints,
                "pagamentos": batch.payment_count,
                "grupos": batch.group_count,
                "candidatos": len(batch.rows),
                "selecionados_previstos": selected,
                "pendencias": warnings,
                "saida_extensao": destination.suffix.lower(),
                "resultado": "excel_gerado",
            },
        )
    except OSError:
        operation_warnings.append(
            "LOG_AUDITORIA_NAO_GRAVADO; o Excel intermediário foi gerado normalmente"
        )
    return PreparationResult(
        destination,
        len(batch.rows),
        selected,
        warnings + len(operation_warnings),
        tuple(operation_warnings),
        batch.payment_count,
        batch.group_count,
        tuple(batch.related_suggestions),
    )


def generate_cnab(
    intermediate_path: str | Path,
    *,
    output_path: str | Path | None = None,
    log_directory: str | Path | None = None,
    lawyer_rules: dict[str, tuple[str, ...]] | None = None,
) -> GenerationResult:
    source = Path(intermediate_path)
    loaded = read_intermediate(source)
    batch = validate_for_generation(loaded, lawyer_rules or {})
    destination = (
        Path(output_path) if output_path else _unique_output(Path.cwd(), "cnab_nox", ".txt")
    )
    write_cnab(batch, destination)
    total_nominal = sum((row["VL_NOMINAL"] for row in batch.rows), Decimal("0.00"))
    total_present = sum((row["VL_PRESENTE"] for row in batch.rows), Decimal("0.00"))
    # Sequences are reserved per PFMI group and may have gaps when not every
    # group is selected yet; the last one used is not first_sequence + count.
    last_sequence = max(int(row["SEU_NUMERO"]) for row in batch.rows)
    result_warnings = list(batch.warnings)
    try:
        _audit(
            log_directory or destination.parent / "logs",
            {
                "evento": "gerar",
                "entrada": _source_fingerprint(source),
                "creditos": len(batch.rows),
                "total_nominal": f"{total_nominal:.2f}",
                "total_presente": f"{total_present:.2f}",
                "sequencia_inicial": batch.first_sequence,
                "sequencia_final": last_sequence,
                "alertas": len(batch.warnings),
                "saida_extensao": destination.suffix.lower(),
                "resultado": "txt_gerado",
            },
        )
    except OSError:
        result_warnings.append("LOG_AUDITORIA_NAO_GRAVADO; o TXT foi gerado normalmente")
    return GenerationResult(
        output_path=destination,
        detail_count=len(batch.rows),
        total_nominal=total_nominal,
        total_present=total_present,
        first_sequence=batch.first_sequence,
        last_sequence=last_sequence,
        warning_count=len(result_warnings),
        warnings=tuple(result_warnings),
    )


def _optional_date(value: Any):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, "DATA_LIQUIDACAO_AUSENTE"
    try:
        return parse_date(value), None
    except ValueError:
        return None, "DATA_LIQUIDACAO_INVALIDA"


def _optional_sequence(value: Any):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, "PRIMEIRA_SEQUENCIA_AUSENTE"
    try:
        return parse_positive_int(value), None
    except ValueError:
        return None, "PRIMEIRA_SEQUENCIA_INVALIDA"


def _unique_output(directory: Path, prefix: str, suffix: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return directory / f"{prefix}_{stamp}{suffix}"


def _source_fingerprint(path: Path) -> dict[str, str]:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return {"extensao": path.suffix.lower(), "sha256_12": digest.hexdigest()[:12]}


def _audit(directory: str | Path, data: dict[str, Any]) -> None:
    log_dir = Path(directory)
    log_dir.mkdir(parents=True, exist_ok=True)
    entry = {"instante": datetime.now().astimezone().isoformat(timespec="seconds"), **data}
    with (log_dir / "gerador_cnab_nox.jsonl").open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, ensure_ascii=False, separators=(",", ":")) + "\n")
