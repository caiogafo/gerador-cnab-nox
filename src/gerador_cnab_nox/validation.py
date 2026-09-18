from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_HALF_UP, Decimal, localcontext
from typing import Any

from .errors import ValidationError
from .matching import CONEXCRED_NAME, is_lawyer_case, lawyer_identity
from .models import CESSAO, FINAL_FIELDS, IMMUTABLE_FIELDS, AnalyticCredit, PreparedRow
from .normalize import (
    MAX_AGGREGATE,
    calculate_commission,
    digits,
    header_key,
    is_formula,
    money,
    normalize_name,
    normalize_name_whitespace,
    parse_date,
    parse_positive_int,
    technical_name,
    validate_document,
)
from .reconciliation import BatchValidationSummary, reconcile_credit, summarize, tolerance_cents
from .stp import SYSTEM_APPROVAL, has_system_approval, system_approval_valid
from .workbook import LoadedIntermediate, _canonical


@dataclass
class ValidatedBatch:
    rows: list[dict[str, Any]]
    liquidation_date: date
    first_sequence: int
    warnings: list[str]
    reconciliation: BatchValidationSummary | None = None


def validate_for_generation(
    loaded: LoadedIntermediate,
    lawyer_rules: dict[str, tuple[str, ...]] | None = None,
    analytic: list[AnalyticCredit] | None = None,
) -> ValidatedBatch:
    lawyer_rules = lawyer_rules or {}
    issues = list(loaded.structural_issues)
    warnings: list[str] = []
    try:
        limit = tolerance_cents()
    except ValueError as exc:
        raise ValidationError([str(exc)], summarize([], other_errors=True)) from exc
    reconciliation_groups: dict[str, dict[str, Any]] = {}
    analytic_by_document: dict[str, list[AnalyticCredit]] = {}
    analytic_by_location: dict[tuple[str, int], list[AnalyticCredit]] = {}
    if analytic:
        for credit in analytic:
            credit_doc = digits(credit.cedent_document)
            if validate_document(credit_doc)[1] is not None:
                analytic_by_document.setdefault(credit_doc, []).append(credit)
            analytic_by_location.setdefault(
                (credit.source_sheet, credit.source_row), []
            ).append(credit)
    used_manual_documents: dict[str, str] = {}
    if issues:
        raise ValidationError(issues, summarize([], other_errors=True))
    requested_groups = {
        row.originals.get("ID_GRUPO") for row in loaded.rows
        if _choice(row.values.get("INCLUIR_CNAB")) == "SIM"
    }
    requested_compositions = {
        row.originals.get("COMPOSICAO_ID") for row in loaded.rows
        if row.originals.get("COMPOSICAO_ID")
        and _choice(row.values.get("INCLUIR_CNAB")) == "SIM"
    }
    active_rows = [
        row for row in loaded.rows if row.originals.get("ID_GRUPO") in requested_groups
        or row.originals.get("COMPOSICAO_ID") in requested_compositions
    ]
    composition_manual_credits = _resolve_composition_manual_selections(
        active_rows, analytic_by_location, issues
    )
    manual_defaults = _manual_credit_defaults(active_rows, composition_manual_credits)
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
    composition_state: dict[str, dict[str, Any]] = {}
    for position, prepared in enumerate(loaded.rows, start=2):
        values = dict(prepared.values)
        line_id = str(values.get("ID_LINHA") or f"linha {position}")
        label = f"CREDITOS linha {position} ({line_id})"
        if any(is_formula(value) for value in values.values()):
            issues.append(f"{label}: fórmulas não são permitidas")
            continue

        choice = _choice(values.get("INCLUIR_CNAB"))
        for control in IMMUTABLE_FIELDS:
            if _canonical(values.get(control)) != _canonical(prepared.originals.get(control)):
                issues.append(f"{label}: controle de proveniência {control} foi alterado")
        if (
            choice == "NAO" and requested_groups
            and prepared.originals.get("ID_GRUPO") not in requested_groups
            and prepared.originals.get("COMPOSICAO_ID") not in requested_compositions
        ):
            continue
        override_key = _override_key(prepared.originals)
        human_override = override_key in composition_manual_credits
        for approval_field in ("APROVADO", "COMPOSICAO_APROVADA", "SELECAO_MANUAL_APROVADA"):
            if _choice(values.get(approval_field)) == SYSTEM_APPROVAL:
                if human_override and prepared.originals.get(approval_field) == SYSTEM_APPROVAL:
                    # A confirmed human location choice supersedes the old robot signature.
                    values[approval_field] = "SIM"
                elif not system_approval_valid(values, prepared.originals, approval_field):
                    issues.append(
                        f"{label}: {approval_field}=SIM_SISTEMA sem origem comprovada "
                        "ou campos aprovados alterados; prepare novamente ou revise manualmente"
                    )
                else:
                    values[approval_field] = "SIM"
                    warnings.append(f"{line_id}: {approval_field}=SIM_SISTEMA (STP auditado)")
        if human_override:
            values.update(manual_defaults.get(line_id, {}))
            if _choice(values.get("APROVADO")) in {"", SYSTEM_APPROVAL}:
                values["APROVADO"] = "SIM"
            warnings.append(f"{line_id}: MANUAL_OVERRIDE=SIM; seleção humana prevalece sobre STP")
        if choice not in {"SIM", "NAO"}:
            issues.append(f"{label}: INCLUIR_CNAB deve ser SIM ou NAO")
        if _choice(values.get("APROVADO")) not in {"SIM", "NAO", ""}:
            issues.append(f"{label}: APROVADO deve ser SIM, NAO ou vazio")
        composicao_aprovada = _choice(values.get("COMPOSICAO_APROVADA"))
        if composicao_aprovada not in {"SIM", "NAO", ""}:
            issues.append(f"{label}: COMPOSICAO_APROVADA deve ser SIM, NAO ou vazio")
        selecao_manual_aprovada = _choice(values.get("SELECAO_MANUAL_APROVADA"))
        if selecao_manual_aprovada not in {"SIM", "NAO", ""}:
            issues.append(f"{label}: SELECAO_MANUAL_APROVADA deve ser SIM, NAO ou vazio")
        # Legacy approval fields remain consistent when filled, but no longer
        # override reconciliation: the configured financial limit is authoritative.
        divergencia_valor_aprovada = _choice(values.get("DIVERGENCIA_VALOR_APROVADA"))
        if divergencia_valor_aprovada not in {"SIM", ""}:
            issues.append(f"{label}: DIVERGENCIA_VALOR_APROVADA deve ser SIM ou vazio")
        divergencia_valor_justificativa = str(
            values.get("DIVERGENCIA_VALOR_JUSTIFICATIVA") or ""
        ).strip()
        if divergencia_valor_aprovada == "SIM" and not divergencia_valor_justificativa:
            issues.append(
                f"{label}: DIVERGENCIA_VALOR_JUSTIFICATIVA é obrigatória quando "
                "DIVERGENCIA_VALOR_APROVADA=SIM"
            )
        if divergencia_valor_justificativa and divergencia_valor_aprovada != "SIM":
            issues.append(
                f"{label}: DIVERGENCIA_VALOR_JUSTIFICATIVA só é aceita quando "
                "DIVERGENCIA_VALOR_APROVADA=SIM"
            )
        composicao_id = str(prepared.originals.get("COMPOSICAO_ID") or "").strip()
        if composicao_id:
            state = composition_state.setdefault(
                composicao_id,
                {
                    "estado": str(prepared.originals.get("COMPOSICAO_ESTADO") or ""),
                    "aprovadas": set(),
                    "escolhas": set(),
                    "labels": [],
                    "div_aprovadas": set(),
                    "div_justificativas": set(),
                    "nominal_sum": Decimal("0.00"),
                    "nominal_count": 0,
                    "nominal_esperado": None,
                    "qtd_componentes": None,
                },
            )
            state["aprovadas"].add(composicao_aprovada)
            state["escolhas"].add(choice)
            state["labels"].append(label)
            state["div_aprovadas"].add(divergencia_valor_aprovada)
            state["div_justificativas"].add(divergencia_valor_justificativa)
            # Soma dos nominais individuais x nominal do crédito compartilhado
            # (regra 5): sempre em Decimal, normalizado em centavos - nunca
            # float. Só o valor esperado é lido aqui (imutável, protegido);
            # a soma real só é fechada depois do laço, com todos os
            # componentes já processados.
            try:
                state["nominal_esperado"] = money(
                    prepared.originals.get("COMPOSICAO_NOMINAL_ANALITICO"), allow_blank=True
                )
            except ValueError:
                state["nominal_esperado"] = None
            if composicao_id in composition_manual_credits:
                state["nominal_esperado"] = composition_manual_credits[composicao_id].nominal_value
            try:
                state["qtd_componentes"] = int(
                    str(prepared.originals.get("COMPOSICAO_QTD_COMPONENTES") or "0")
                )
            except ValueError:
                state["qtd_componentes"] = None
            divergencia_valor_ok = (
                state["estado"] == "BLOQUEADA_DIVERGENCIA_VALOR"
                and _within_tolerance(
                    prepared.originals.get("COMPOSICAO_TOTAL_PFMI"),
                    prepared.originals.get("COMPOSICAO_TOTAL_ANALITICO"), limit,
                )
            )
            effectively_proposta = (
                state["estado"] == "PROPOSTA"
                or composicao_id in composition_manual_credits
                or divergencia_valor_ok
            )
            if composicao_aprovada == "SIM" and not effectively_proposta:
                issues.append(
                    f"{label}: COMPOSICAO_APROVADA=SIM só é permitido quando "
                    "COMPOSICAO_ESTADO é PROPOSTA, a seleção manual do principal fechou "
                    "com um crédito identificado, ou a divergência está dentro da tolerância"
                )
            if choice == "SIM" and composicao_aprovada != "SIM":
                issues.append(
                    f"{label}: título de composição só pode ser incluído (INCLUIR_CNAB=SIM) "
                    "quando COMPOSICAO_APROVADA=SIM para todo o grupo"
                )
            if choice == "SIM" and values.get("VL_NOMINAL") in (None, ""):
                issues.append(f"{label}: PREENCHER_VL_NOMINAL_COMPOSICAO")
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
        manual_credit = None
        if (
            human_override
            or not analytic_line or not analytic_reference or not credit_id
        ):
            if human_override:
                # Principal de composição resolvido por referência manual
                # única (ver _resolve_composition_manual_selections): todo
                # componente (principal e satélites) compartilha este mesmo
                # crédito, igual à composição detectada automaticamente.
                manual_credit = composition_manual_credits[override_key]
            else:
                manual_credit = _resolve_manual_selection(
                    values,
                    prepared.originals,
                    target,
                    selecao_manual_aprovada,
                    analytic_by_document,
                    label,
                    issues,
                )
                if manual_credit is None:
                    continue
                manual_doc = digits(values.get("SELECAO_MANUAL_DOCUMENTO"))
                previous_owner = used_manual_documents.get(manual_doc)
                if previous_owner is not None and previous_owner != group_id:
                    issues.append(
                        f"{label}: o mesmo documento de seleção manual já foi usado em outro grupo"
                    )
                    continue
                used_manual_documents[manual_doc] = group_id
            credit_id = manual_credit.credit_id
            analytic_line = str(manual_credit.source_row)
            analytic_reference = manual_credit.reference
            values["ID_CREDITO"] = credit_id
            values["LINHA_ANALITICO"] = analytic_line
            values["REFERENCIA_ANALITICO"] = analytic_reference
        # Reconcile values only after the credit identity is resolved.
        identity = credit_id
        # Componentes de uma mesma composição compartilham deliberadamente um
        # único crédito (ver COMPOSICAO_ID); só é reutilização indevida se o
        # crédito aparecer fora dessa composição especifica.
        owner = f"COMPOSICAO:{composicao_id}" if composicao_id else group_id
        existing = selected_credit_identities.get(identity)
        if existing is not None and existing != owner:
            issues.append(f"{label}: o mesmo crédito do Analítico foi selecionado mais de uma vez")
            continue
        selected_credit_identities[identity] = owner
        try:
            if manual_credit is not None:
                _, _, analytic_value = calculate_commission(
                    manual_credit.nominal_value, manual_credit.acquisition_value,
                    Decimal(str(prepared.originals.get("PERCENTUAL_USADO") or "0")),
                )
            else:
                analytic_raw = (
                    prepared.originals.get("COMPOSICAO_TOTAL_ANALITICO")
                    if composicao_id else prepared.originals.get("DIVERGENCIA_VALOR_ANALITICO")
                )
                if analytic_raw in (None, "") and not composicao_id:
                    analytic_raw = prepared.originals.get("VL_PRESENTE")
                analytic_value = money(analytic_raw, maximum=MAX_AGGREGATE)
            if analytic_value is None or target is None:
                raise ValueError("valor de origem ausente")
            entry = reconciliation_groups.setdefault(owner, {
                "targets": {}, "credits": {}, "rows": [],
                "document": manual_credit.cedent_document if manual_credit else (
                    prepared.originals.get("DOC_CEDENTE_ANALITICO") or values.get("DOC_CEDENTE")
                ),
                "name": manual_credit.cedent_name if manual_credit else (
                    prepared.originals.get("NOME_CEDENTE_ANALITICO") or values.get("NOME_CEDENTE")
                ),
            })
            entry["targets"][group_id] = target
            if identity in entry["credits"] and entry["credits"][identity] != analytic_value:
                raise ValueError("valor do crédito inconsistente")
            entry["credits"][identity] = analytic_value
        except (ValueError, ArithmeticError) as exc:
            issues.append(f"{label}: reconciliação indisponível ({exc})")
            continue
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
            if manual_credit is not None:
                # Resolved output only; the workbook's protected snapshot stays intact.
                parsed["ID_CREDITO"] = credit_id
                parsed["ABA_ANALITICO"] = manual_credit.source_sheet
                parsed["LINHA_ANALITICO"] = manual_credit.source_row
                parsed["REFERENCIA_ANALITICO"] = analytic_reference
                parsed["HASH_ANALITICO"] = manual_credit.source_hash
            entry["rows"].append((group_id, parsed))
            group["sum"] += parsed["VL_PRESENTE"]
            group["count"] += 1
            selected.append((reserved_sequence, parsed))
            if composicao_id:
                composition_state[composicao_id]["nominal_sum"] += parsed["VL_NOMINAL"]
                composition_state[composicao_id]["nominal_count"] += 1

    reconciliation_records = []
    for owner, entry in reconciliation_groups.items():
        try:
            record = reconcile_credit(
                owner, entry["document"], entry["name"],
                sum(entry["targets"].values(), Decimal(0)),
                sum(entry["credits"].values(), Decimal(0)), limit_cents=limit,
            )
        except ValueError as exc:
            issues.append(f"{owner}: reconciliação indisponível ({exc})")
            continue
        reconciliation_records.append(record)
        if record.errorMessage:
            issues.append(record.errorMessage)
        elif record.warningMessage:
            warnings.append(record.warningMessage)
            for group_id, parsed in entry["rows"]:
                if sum(g == group_id for g, _ in entry["rows"]) != 1:
                    issues.append(
                        f"{group_id}: tolerância não define rateio entre múltiplos créditos"
                    )
                    continue
                groups[group_id]["sum"] += entry["targets"][group_id] - parsed["VL_PRESENTE"]
                parsed["VL_PRESENTE"] = entry["targets"][group_id]

    for composicao_id, state in composition_state.items():
        if len(state["aprovadas"]) > 1:
            issues.append(
                f"Composição {composicao_id}: COMPOSICAO_APROVADA deve ser igual em "
                f"todas as linhas ({', '.join(state['labels'])})"
            )
        if len(state["escolhas"]) > 1:
            issues.append(
                f"Composição {composicao_id}: INCLUIR_CNAB deve ser igual em todas as "
                f"linhas, tudo ou nada ({', '.join(state['labels'])})"
            )
        if len(state["div_aprovadas"]) > 1:
            issues.append(
                f"Composição {composicao_id}: DIVERGENCIA_VALOR_APROVADA deve ser igual "
                f"em todas as linhas ({', '.join(state['labels'])})"
            )
        if len(state["div_justificativas"]) > 1:
            issues.append(
                f"Composição {composicao_id}: DIVERGENCIA_VALOR_JUSTIFICATIVA deve ser "
                f"igual em todas as linhas ({', '.join(state['labels'])})"
            )
        # Soma dos nominais individuais x nominal do crédito compartilhado
        # (regra adicional de segurança, não é afrouxamento): só avalia
        # quando a composição inteira foi incluída e todo componente já
        # tem VL_NOMINAL validado - do contrário outras checagens já
        # bloqueiam a linha/grupo em falta, sem duplicar o motivo. Decimal
        # exato, sem tolerância: qualquer diferença, inclusive R$0,01,
        # bloqueia.
        if (
            state["escolhas"] == {"SIM"}
            and state["qtd_componentes"] is not None
            and state["nominal_count"] == state["qtd_componentes"]
            and state["nominal_esperado"] is not None
        ):
            diferenca = state["nominal_sum"] - state["nominal_esperado"]
            if diferenca != Decimal("0.00"):
                issues.append(
                    f"Composição {composicao_id}: soma dos nominais informados "
                    f"({state['nominal_sum']:.2f}) diverge do nominal do crédito "
                    f"compartilhado ({state['nominal_esperado']:.2f}); diferença de "
                    f"{diferenca:+.2f}"
                )

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
        raise ValidationError(
            list(dict.fromkeys(issues)), summarize(reconciliation_records, other_errors=True)
        )

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
        raise ValidationError(issues, summarize(reconciliation_records, other_errors=True))
    return ValidatedBatch(
        rows=[row for _, row in selected],
        liquidation_date=liquidation_date,
        first_sequence=first_sequence,
        warnings=list(dict.fromkeys(warnings)),
        reconciliation=summarize(reconciliation_records),
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

    cedent_name = normalize_name_whitespace(values.get("NOME_CEDENTE"))
    parsed["NOME_CEDENTE"] = cedent_name
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


def _override_key(originals):
    return str(originals.get("COMPOSICAO_ID") or "").strip() or (
        "GRUPO:" + str(originals.get("ID_GRUPO") or "")
    )


def _manual_credit_defaults(rows, overrides):
    """Refresh obsolete automatic defaults, preserving every human-edited final field."""
    grouped = {}
    for row in rows:
        key = _override_key(row.originals)
        if key in overrides:
            grouped.setdefault(key, []).append(row)
    result = {}
    for key, components in grouped.items():
        credit = overrides[key]
        composition = bool(components[0].originals.get("COMPOSICAO_ID"))
        principal = str(components[0].originals.get("COMPOSICAO_PARTICIPANTES") or "")
        principal = principal.split(";")[0].strip()
        nominals = None
        if credit.nominal_value is not None and all(
            _canonical(r.values.get("VL_NOMINAL")) == _canonical(r.originals.get("VL_NOMINAL"))
            for r in components
        ):
            try:
                amounts = [
                    money(r.originals.get("TOTAL_PFMI_GRUPO"), maximum=MAX_AGGREGATE)
                    for r in components
                ]
                with localcontext() as context:
                    context.prec = 80
                    total = sum(amounts, Decimal(0))
                    if total > 0:
                        nominals = [
                            (credit.nominal_value * amount / total).quantize(
                                Decimal("0.01"), rounding=ROUND_HALF_UP,
                            ) for amount in amounts
                        ]
                        principal_index = next((
                            i for i, r in enumerate(components)
                            if r.originals.get("NOME_CEDENTE_PFMI") == principal
                        ), 0)
                        nominals[principal_index] += credit.nominal_value - sum(nominals)
            except (ValueError, TypeError, ArithmeticError):
                pass  # Existing validation reports invalid source amounts.
        for index, row in enumerate(components):
            defaults = {}
            if not composition or row.originals.get("NOME_CEDENTE_PFMI") == principal:
                defaults.update({
                    "NOME_CEDENTE": credit.cedent_name,
                    "DOC_CEDENTE": digits(credit.cedent_document),
                    "TIPO_PESSOA_CEDENTE": validate_document(credit.cedent_document)[1],
                })
            if row.originals.get("MODALIDADE") != CESSAO:
                defaults["DT_EMISSAO_TITULO"] = credit.signature_date
            if nominals is not None:
                defaults["VL_NOMINAL"] = nominals[index]
            defaults["VL_PRESENTE"] = row.originals.get("TOTAL_PFMI_GRUPO")
            result[str(row.values.get("ID_LINHA"))] = {
                field: value for field, value in defaults.items()
                if _canonical(row.values.get(field)) == _canonical(row.originals.get(field))
            }
    return result


def _resolve_composition_manual_selections(
    rows: list[PreparedRow],
    analytic_by_location: dict[tuple[str, int], list[AnalyticCredit]],
    issues: list[str],
) -> dict[str, AnalyticCredit]:
    """Explicit, human-initiated fallback for a composition principal that the
    automatic detection could not locate (COMPOSICAO_ESTADO != PROPOSTA).

    Identifies a SINGLE physical record of the Analítico - its own sheet
    (ABA) and physical row (LINHA), the same coordinates that already make
    AnalyticCredit.credit_id deterministic and unique - never a shared
    campaign/batch field such as "Prospect", which some real Analítico files
    reuse across hundreds of unrelated credits and can never safely identify
    one of them. COMPOSICAO_CANDIDATOS (matching.py) lists exact-value
    matches (name, masked document, value, aba/linha) for the operator to
    review before choosing.

    Mirrors _resolve_manual_selection's safety properties at the composition
    level: never triggered unless the operator filled both ABA and LINHA
    after checking the physical instrument; the location must resolve to
    exactly one credit and be identical on every component row (auditable -
    no silent per-row override, and no tampering by pointing only the
    satellite somewhere else). Financial tolerance is checked centrally
    after resolution, against the combined PFMI total and one shared credit.
    Approving the name divergence (APROVADO) and the
    composition itself (COMPOSICAO_APROVADA) still goes through the existing
    fields, exactly like an automatically detected PROPOSTA composition.
    """
    groups: dict[str, dict[str, Any]] = {}
    for prepared in rows:
        if not prepared.originals.get("COMPOSICAO_ID") and _choice(
            prepared.values.get("INCLUIR_CNAB")
        ) != "SIM":
            continue
        composicao_id = _override_key(prepared.originals)
        snapshot_location = bool(
            prepared.originals.get("ID_CREDITO")
            and has_system_approval(prepared.originals, "SELECAO_MANUAL_APROVADA")
            and system_approval_valid(
                prepared.values, prepared.originals, "SELECAO_MANUAL_APROVADA",
            )
        )
        entry = groups.setdefault(
            composicao_id,
            {
                "localizacoes": set(),
                "principal": str(prepared.originals.get("COMPOSICAO_PARTICIPANTES") or "")
                .split(";")[0]
                .strip() or prepared.originals.get("NOME_CEDENTE_PFMI"),
                "rows": [],
                "approvals": set(),
                "snapshot_locations": [],
            },
        )
        aba_raw = str(prepared.values.get("COMPOSICAO_SELECAO_MANUAL_ABA") or "").strip()
        linha_raw = prepared.values.get("COMPOSICAO_SELECAO_MANUAL_LINHA")
        linha_raw = "" if linha_raw in (None, "") else str(linha_raw).strip()
        entry["localizacoes"].add((aba_raw, linha_raw))
        human_signature = (
            _choice(prepared.values.get("SELECAO_MANUAL_APROVADA")) == "SIM"
            or _choice(prepared.values.get(
                "COMPOSICAO_APROVADA" if prepared.originals.get("COMPOSICAO_ID") else "APROVADO"
            )) == "SIM"
        )
        entry["snapshot_locations"].append(snapshot_location)
        entry["approvals"].add(human_signature)
        entry["rows"].append(
            {
                "name": str(prepared.originals.get("NOME_CEDENTE_PFMI") or "").strip(),
                "target": prepared.originals.get("TOTAL_PFMI_GRUPO"),
                "percentual": prepared.originals.get("PERCENTUAL_USADO"),
            }
        )

    resolved: dict[str, AnalyticCredit] = {}
    for composicao_id, entry in groups.items():
        if all(entry["snapshot_locations"]):
            # Unchanged system selections remain usable without reopening source files.
            continue
        attempted = {loc for loc in entry["localizacoes"] if loc != ("", "")}
        if not attempted:
            continue  # operador não tentou seleção manual; composição segue bloqueada
        label = f"Composição {composicao_id} (seleção manual do principal)"
        if len(entry["localizacoes"]) > 1:
            issues.append(
                f"{label}: COMPOSICAO_SELECAO_MANUAL_ABA/LINHA devem ser preenchidos e "
                "iguais em todas as linhas da composição"
            )
            continue
        aba_raw, linha_raw = next(iter(attempted))
        linha = _int_or_none(linha_raw)
        if not aba_raw or linha is None:
            issues.append(
                f"{label}: informe COMPOSICAO_SELECAO_MANUAL_ABA e "
                "COMPOSICAO_SELECAO_MANUAL_LINHA (os dois) para usar a seleção manual"
            )
            continue
        matches = analytic_by_location.get((aba_raw, linha), [])
        if len(matches) == 0:
            issues.append(
                f"{label}: nenhum crédito do Analítico foi encontrado em "
                f"{aba_raw!r}, linha {linha}"
            )
            continue
        if len(matches) > 1:
            issues.append(
                f"{label}: mais de um crédito do Analítico corresponde a "
                f"{aba_raw!r}, linha {linha}"
            )
            continue
        credit = matches[0]
        if entry["approvals"] != {True}:
            issues.append(f"{label}: confirme a seleção humana com SIM em todas as linhas")
            continue
        principal_rows = [r for r in entry["rows"] if r["name"] == entry["principal"]]
        if len(principal_rows) != 1:
            issues.append(
                f"{label}: não foi possível identificar de forma única o principal da composição"
            )
            continue
        targets: list[Decimal | None] = []
        for r in entry["rows"]:
            try:
                targets.append(money(r["target"], maximum=MAX_AGGREGATE))
            except ValueError:
                targets.append(None)
        if any(t is None for t in targets):
            issues.append(f"{label}: total PFMI original inválido em algum componente")
            continue
        try:
            rate = Decimal(str(principal_rows[0]["percentual"] or "0"))
            _, _, expected = calculate_commission(
                credit.nominal_value, credit.acquisition_value, rate
            )
        except (ValueError, ArithmeticError) as exc:
            issues.append(f"{label}: {exc}")
            continue
        if expected is None:
            issues.append(
                f"{label}: não foi possível calcular a aquisição esperada do crédito selecionado"
            )
            continue
        # Financial tolerance is evaluated centrally after identity resolution.
        resolved[composicao_id] = credit
    return resolved


def _resolve_manual_selection(
    values: dict[str, Any],
    originals: dict[str, Any],
    target: Decimal | None,
    selecao_manual_aprovada: str,
    analytic_by_document: dict[str, list[AnalyticCredit]],
    label: str,
    issues: list[str],
) -> AnalyticCredit | None:
    """Explicit, human-initiated fallback for a row with no automatic credit.

    Never triggered automatically: the operator must have already read the
    physical instrument, typed the credit's own CPF/CNPJ, and set
    SELECAO_MANUAL_APROVADA=SIM. A single valid, exact document match is
    required (no similarity, never across a different falência - the
    document must belong to a credit already filtered by analytic_by_document
    at the caller). Financial closure reuses the existing homologated
    formula in Decimal; financial tolerance is checked centrally after resolution.
    """
    manual_doc_raw = values.get("SELECAO_MANUAL_DOCUMENTO")
    manual_doc = digits(manual_doc_raw) if manual_doc_raw else ""
    if not manual_doc:
        issues.append(
            f"{label}: linha sem crédito original do Analítico não pode ser selecionada"
        )
        return None
    if validate_document(manual_doc)[1] is None:
        issues.append(f"{label}: SELECAO_MANUAL_DOCUMENTO inválido")
        return None
    if selecao_manual_aprovada != "SIM":
        issues.append(
            f"{label}: SELECAO_MANUAL_APROVADA deve ser SIM para usar a seleção manual "
            "por documento"
        )
        return None
    matches = analytic_by_document.get(manual_doc, [])
    if len(matches) == 0:
        issues.append(
            f"{label}: nenhum crédito do Analítico tem o documento de SELECAO_MANUAL_DOCUMENTO"
        )
        return None
    if len(matches) > 1:
        issues.append(
            f"{label}: mais de um crédito do Analítico tem o documento de "
            "SELECAO_MANUAL_DOCUMENTO"
        )
        return None
    credit = matches[0]
    if target is None:
        issues.append(f"{label}: total PFMI original inválido para conferir a seleção manual")
        return None
    try:
        rate = Decimal(str(originals.get("PERCENTUAL_USADO") or "0"))
        _, _, expected = calculate_commission(credit.nominal_value, credit.acquisition_value, rate)
    except (ValueError, ArithmeticError) as exc:
        issues.append(f"{label}: {exc}")
        return None
    if expected is None:
        issues.append(
            f"{label}: não foi possível calcular a aquisição esperada do crédito selecionado"
        )
        return None
    # Financial tolerance is evaluated centrally after identity resolution.
    entered_nominal = values.get("VL_NOMINAL")
    try:
        if entered_nominal is None or money(entered_nominal) != credit.nominal_value:
            raise ValueError
    except ValueError:
        issues.append(
            f"{label}: VL_NOMINAL deve ser exatamente o nominal do crédito selecionado "
            "manualmente"
        )
        return None
    return credit


def _choice(value: Any) -> str:
    key = header_key(value)
    return {"SIM": "SIM", "NAO": "NAO", "SIMSISTEMA": SYSTEM_APPROVAL}.get(key, key)


def _within_tolerance(pfmi: Any, analytic: Any, limit: int) -> bool:
    try:
        result = reconcile_credit(
            "", "", "", money(pfmi, maximum=MAX_AGGREGATE),
            money(analytic, maximum=MAX_AGGREGATE), limit_cents=limit,
        )
        return result.status != "CRITICAL_ERROR"
    except (ValueError, AttributeError):
        return False


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
