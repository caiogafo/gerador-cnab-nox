from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Any

from .errors import ValidationError
from .matching import CONEXCRED_NAME, is_lawyer_case, lawyer_identity
from .models import CESSAO, FINAL_FIELDS, IMMUTABLE_FIELDS
from .normalize import (
    MAX_AGGREGATE,
    header_key,
    is_formula,
    money,
    normalize_name,
    parse_date,
    parse_positive_int,
    technical_name,
    validate_document,
)
from .workbook import LoadedIntermediate, _canonical


@dataclass
class ValidatedBatch:
    rows: list[dict[str, Any]]
    liquidation_date: date
    first_sequence: int
    warnings: list[str]


def validate_for_generation(
    loaded: LoadedIntermediate, lawyer_rules: dict[str, tuple[str, ...]] | None = None
) -> ValidatedBatch:
    lawyer_rules = lawyer_rules or {}
    issues = list(loaded.structural_issues)
    warnings: list[str] = []
    if issues:
        raise ValidationError(issues)
    if loaded.unknown_count:
        issues.append(
            f"Existem {loaded.unknown_count} linha(s) não reconhecida(s) na PFMI; "
            "corrija e prepare novamente"
        )
    try:
        liquidation_date = parse_date(loaded.liquidation_date)
    except ValueError as exc:
        liquidation_date = None
        issues.append(f"DATA_LIQUIDACAO: {exc}")
    try:
        first_sequence = parse_positive_int(loaded.first_sequence)
    except ValueError as exc:
        first_sequence = None
        issues.append(f"PRIMEIRA_SEQUENCIA: {exc}")

    selected: list[tuple[int | None, dict[str, Any]]] = []
    groups: dict[str, dict[str, Any]] = {}
    selected_credit_identities: dict[str, str] = {}
    for position, prepared in enumerate(loaded.rows, start=2):
        values = dict(prepared.values)
        line_id = str(values.get("ID_LINHA") or f"linha {position}")
        label = f"CREDITOS linha {position} ({line_id})"
        if any(is_formula(value) for value in values.values()):
            issues.append(f"{label}: fórmulas não são permitidas")
            continue

        choice = _choice(values.get("INCLUIR_CNAB"))
        if choice not in {"SIM", "NAO"}:
            issues.append(f"{label}: INCLUIR_CNAB deve ser SIM ou NAO")
        if _choice(values.get("APROVADO")) not in {"SIM", "NAO", ""}:
            issues.append(f"{label}: APROVADO deve ser SIM, NAO ou vazio")
        for control in IMMUTABLE_FIELDS:
            if _canonical(values.get(control)) != _canonical(prepared.originals.get(control)):
                issues.append(f"{label}: controle de proveniência {control} foi alterado")
        parameter_issue = prepared.originals.get("PENDENCIAS_PARAMETROS")
        if parameter_issue:
            issues.append(f"{label}: {parameter_issue}; corrija o parâmetro e prepare novamente")
        source_issues = str(prepared.originals.get("PENDENCIAS_FONTE_PFMI") or "").strip()
        if source_issues:
            issues.append(
                f"{label}: pendência de origem PFMI ({source_issues}); "
                "corrija a PFMI e prepare novamente"
            )
        group_id = str(prepared.originals.get("ID_GRUPO") or "").strip()
        try:
            target = money(prepared.originals.get("TOTAL_PFMI_GRUPO"), maximum=MAX_AGGREGATE)
        except ValueError:
            target = None
            issues.append(f"{label}: total PFMI original inválido")
        group = groups.setdefault(group_id, {"target": target, "sum": Decimal("0.00"), "count": 0})
        if group["target"] != target:
            issues.append(f"{label}: total PFMI inconsistente dentro do grupo")

        if choice != "SIM":
            continue
        analytic_line = str(prepared.originals.get("LINHA_ANALITICO") or "").strip()
        analytic_reference = str(prepared.originals.get("REFERENCIA_ANALITICO") or "").strip()
        credit_id = str(prepared.originals.get("ID_CREDITO") or "")
        if not analytic_line or not analytic_reference or not credit_id:
            issues.append(
                f"{label}: linha sem crédito original do Analítico não pode ser selecionada"
            )
            continue
        identity = credit_id
        if identity in selected_credit_identities:
            issues.append(f"{label}: o mesmo crédito do Analítico foi selecionado mais de uma vez")
            continue
        selected_credit_identities[identity] = group_id
        raw_reserved = prepared.originals.get("SEU_NUMERO")
        if raw_reserved in (None, ""):
            # Prepared before a first sequence was available (e.g. informed
            # later directly in RESUMO): nothing was reserved per group yet.
            reserved_sequence = None
        else:
            try:
                reserved_sequence = parse_positive_int(raw_reserved)
            except ValueError:
                reserved_sequence = None
                issues.append(
                    f"{label}: sequência reservada do grupo é inválida; prepare novamente"
                )
        parsed = _validate_selected(
            values, prepared.originals, label, issues, warnings, lawyer_rules
        )
        if parsed is not None:
            group["sum"] += parsed["VL_PRESENTE"]
            group["count"] += 1
            selected.append((reserved_sequence, parsed))

    for group_id, group in groups.items():
        if not group_id:
            issues.append("Existe linha sem ID_GRUPO")
            continue
        if group["count"] == 0:
            issues.append(f"Grupo {group_id}: nenhum crédito selecionado")
        elif group["target"] is not None and group["sum"] != group["target"]:
            difference = group["sum"] - group["target"]
            issues.append(
                f"Grupo {group_id}: DIVERGENCIA_VALOR de {difference:+.2f}; "
                "a soma de VL_PRESENTE deve fechar exatamente com a PFMI"
            )

    if all(g["target"] is not None for g in groups.values()) and sum(
        (g["sum"] for g in groups.values()), Decimal(0)
    ) != sum((g["target"] for g in groups.values()), Decimal(0)):
        issues.append("Lote: DIVERGENCIA_VALOR; aquisição deve fechar exatamente com a PFMI")
    reserved = [s for s, _ in selected if s is not None]
    if (first_sequence is not None and first_sequence > 9_999_999_999) or (
        reserved and max(reserved) > 9_999_999_999
    ):
        issues.append("Faixa de sequência não cabe em NU_DOCUMENTO (10 posições)")
    if len(selected) + 2 > 999_999:
        issues.append("Quantidade de registros não cabe na sequência física do CNAB")
    if not selected:
        issues.append("Nenhum crédito válido foi selecionado")
    if issues:
        raise ValidationError(list(dict.fromkeys(issues)))

    assert liquidation_date is not None
    assert first_sequence is not None
    if all(s is None for s, _ in selected):
        # Nothing was reserved at prepare time (first sequence was supplied
        # later, directly in RESUMO): assign sequentially now, in selection
        # order, same as before group reservation existed.
        selected = [
            (first_sequence + offset, row) for offset, (_, row) in enumerate(selected)
        ]
    used_sequences: dict[int, str] = {}
    for reserved_sequence, row in selected:
        if reserved_sequence in used_sequences:
            issues.append(
                f"Sequência {reserved_sequence} duplicada entre grupos selecionados "
                f"({used_sequences[reserved_sequence]} e {row['ID_LINHA']})"
            )
        else:
            used_sequences[reserved_sequence] = row["ID_LINHA"]
        if str(row.get("SEU_NUMERO") or "").strip() not in {"", str(reserved_sequence)}:
            warnings.append(f"{row['ID_LINHA']}: SEU_NUMERO editado foi recalculado")
        if str(row.get("NU_DOCUMENTO") or "").strip() not in {"", str(reserved_sequence)}:
            warnings.append(f"{row['ID_LINHA']}: NU_DOCUMENTO editado foi recalculado")
        row["SEU_NUMERO"] = reserved_sequence
        row["NU_DOCUMENTO"] = reserved_sequence
        if row["DT_EMISSAO_TITULO"] > liquidation_date:
            issues.append(f"{row['ID_LINHA']}: DT_EMISSAO_TITULO é posterior à DATA_LIQUIDACAO")
    if issues:
        raise ValidationError(issues)
    return ValidatedBatch(
        rows=[row for _, row in selected],
        liquidation_date=liquidation_date,
        first_sequence=first_sequence,
        warnings=list(dict.fromkeys(warnings)),
    )


def _validate_selected(
    values: dict[str, Any],
    originals: dict[str, Any],
    label: str,
    issues: list[str],
    warnings: list[str],
    lawyer_rules: dict[str, tuple[str, ...]],
) -> dict[str, Any] | None:
    parsed = dict(values)
    start_issues = len(issues)
    cedent_doc, cedent_type = validate_document(values.get("DOC_CEDENTE"))
    debtor_doc, debtor_type = validate_document(values.get("DOC_SACADO"))
    if cedent_type is None:
        issues.append(f"{label}: DOC_CEDENTE inválido")
    if debtor_type is None:
        issues.append(f"{label}: DOC_SACADO inválido")
    parsed["DOC_CEDENTE"] = cedent_doc
    parsed["DOC_SACADO"] = debtor_doc
    if cedent_type is not None and _int_or_none(values.get("TIPO_PESSOA_CEDENTE")) != cedent_type:
        issues.append(f"{label}: TIPO_PESSOA_CEDENTE incompatível com o documento")
    if debtor_type is not None and _int_or_none(values.get("TIPO_PESSOA_SACADO")) != debtor_type:
        issues.append(f"{label}: TIPO_PESSOA_SACADO incompatível com o documento")
    parsed["TIPO_PESSOA_CEDENTE"] = cedent_type
    parsed["TIPO_PESSOA_SACADO"] = debtor_type

    for field in ("VL_NOMINAL", "VL_PRESENTE", "VALOR_PAGO_TITULO"):
        try:
            amount = money(values.get(field))
            assert amount is not None
            if amount < 0:
                raise ValueError("valor negativo")
            parsed[field] = amount
        except (ValueError, AssertionError) as exc:
            issues.append(f"{label}: {field} inválido ({exc})")
    if parsed.get("VALOR_PAGO_TITULO") not in (None, Decimal("0.00")):
        issues.append(f"{label}: VALOR_PAGO_TITULO deve ser zero para cessão nova")
    for field in ("DT_VENCIMENTO", "DT_EMISSAO_TITULO"):
        try:
            parsed[field] = parse_date(values.get(field))
        except ValueError as exc:
            issues.append(f"{label}: {field} inválida ({exc})")

    for field, expected in (("TP_TITULO", 24), ("COOBRIGACAO", 2), ("MOVIMENTO", 1)):
        if _int_or_none(values.get(field)) != expected:
            issues.append(f"{label}: {field} deve ser {expected}")
        parsed[field] = expected

    cedent_name = str(values.get("NOME_CEDENTE") or "").strip()
    debtor_name = str(values.get("NOME_SACADO") or "").strip()
    for field, value in (("NOME_CEDENTE", cedent_name), ("NOME_SACADO", debtor_name)):
        try:
            technical = technical_name(value)
            parsed[field + "_TECNICO"] = technical[:40]
            if len(technical) > 40:
                warnings.append(f"{label}: {field} NOME_TRUNCADO_CNAB (40 bytes)")
        except ValueError as exc:
            issues.append(f"{label}: {field} {exc}")
    expected_name = originals.get("NOME_ESPERADO") or originals.get("NOME_CEDENTE_PFMI")
    modality = originals.get("MODALIDADE")
    if modality == CESSAO:
        if cedent_type != 2:
            issues.append(f"{label}: DOC_CEDENTE da Conexcred deve ser CNPJ válido")
        if parsed.get("NOME_CEDENTE_TECNICO") != CONEXCRED_NAME:
            issues.append(f"{label}: cedente da Cessão da Cessão deve ser somente {CONEXCRED_NAME}")
    failure_reference = originals.get("FALENCIA_PFMI")
    if modality != CESSAO and (
        is_lawyer_case(failure_reference, lawyer_rules)
        or is_lawyer_case(originals.get("FALENCIA_CANONICA"), lawyer_rules)
    ):
        reference_identity = lawyer_identity(expected_name, failure_reference, lawyer_rules)
        final_identity = lawyer_identity(cedent_name, failure_reference, lawyer_rules)
        original_creditors = [
            normalize_name(part)
            for field in ("NOME_CEDENTE_PFMI", "NOME_CEDENTE_ANALITICO")
            for part in re.findall(r"\(([^()]+)\)", str(originals.get(field) or ""))
        ]
        composed = any(
            part and f" {part} " in f" {normalize_name(cedent_name)} "
            for part in original_creditors
        )
        if not final_identity or "(" in cedent_name or ")" in cedent_name or composed:
            issues.append(f"{label}: Papéis Independência exige somente o advogado, sem composição")
        if reference_identity and final_identity != reference_identity:
            issues.append(f"{label}: advogado final difere da identidade indicada na fonte")
    if (
        normalize_name(cedent_name) != normalize_name(expected_name)
        and _choice(values.get("APROVADO")) != "SIM"
    ):
        issues.append(f"{label}: DIVERGENCIA_NOME; corrija NOME_CEDENTE ou marque APROVADO=SIM")

    address = str(values.get("ENDERECO") or "").strip()
    if len(address) > 40:
        issues.append(f"{label}: ENDERECO excede 40 caracteres")
    cep = str(values.get("CEP") or "").strip()
    if cep and (not cep.isdigit() or len(cep) > 8):
        issues.append(f"{label}: CEP deve conter no máximo 8 dígitos")
    nfe = str(values.get("NFE") or "").strip()
    if nfe and (not nfe.isdigit() or len(nfe) > 44):
        issues.append(f"{label}: NFE deve conter no máximo 44 dígitos")
    indexer = str(values.get("INDEXADOR") or "").strip()
    if indexer:
        issues.append(f"{label}: INDEXADOR deve ficar vazio na V1")
    rate_value = values.get("TAXA_INDEXADOR")
    if rate_value not in (None, ""):
        try:
            rate = Decimal(str(rate_value).replace(",", "."))
            if rate < 0:
                raise ValueError("taxa negativa")
            issues.append(f"{label}: TAXA_INDEXADOR deve ficar vazia na V1")
            parsed["TAXA_INDEXADOR"] = rate
        except Exception as exc:
            issues.append(f"{label}: TAXA_INDEXADOR inválida ({exc})")
    else:
        parsed["TAXA_INDEXADOR"] = None

    for field in ("ENDERECO", "INDEXADOR", "NFE"):
        try:
            text = str(values.get(field) or "")
            if any(ord(char) < 32 or 127 <= ord(char) < 160 for char in text):
                issues.append(f"{label}: {field} possui caractere de controle")
            text.encode("cp1252")
        except UnicodeEncodeError:
            issues.append(f"{label}: {field} possui caractere incompatível com Windows-1252")

    for field in FINAL_FIELDS:
        if field in {"SEU_NUMERO", "NU_DOCUMENTO"}:
            continue
        if _comparable(values.get(field)) != _comparable(originals.get(field)):
            warnings.append(f"{values.get('ID_LINHA')}: {field} foi alterado no Excel")
    return parsed if len(issues) == start_issues else None


def _choice(value: Any) -> str:
    key = header_key(value)
    return {"SIM": "SIM", "NAO": "NAO"}.get(key, key)


def _int_or_none(value: Any) -> int | None:
    try:
        numeric = Decimal(str(value))
        if numeric != numeric.to_integral_value():
            return None
        return int(numeric)
    except (TypeError, ValueError, ArithmeticError):
        return None


def _comparable(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        return str(Decimal(str(value)).normalize())
    return str(value).strip()
