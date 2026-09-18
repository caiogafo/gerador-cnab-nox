"""Offline STP evidence rules. Monetary coincidence alone never proves identity."""

from __future__ import annotations

import json
import re
from collections import defaultdict
from difflib import SequenceMatcher

from .models import CESSAO
from .normalize import digits, normalize_name, validate_document
from .reconciliation import reconcile_credit, tolerance_cents

SYSTEM_APPROVAL = "SIM_SISTEMA"


def identity_tokens(name):
    ignored = {"ESPOLIO", "DE", "DA", "DO", "DAS", "DOS", "E", "LTDA", "LIMITADA"}
    return [t for t in normalize_name(name).split() if t not in ignored and not t.isdigit()]


def strong_name_match(left, right):
    a, b = identity_tokens(left), identity_tokens(right)
    if min(len(a), len(b)) < 2:
        return False
    if a == b:
        return True
    # An abbreviated name must contain at least two complete leading tokens.
    shorter, longer = sorted((a, b), key=len)
    if longer[:len(shorter)] == shorter and len(shorter) / len(longer) >= 0.5:
        return True
    return a[0] == b[0] and SequenceMatcher(None, " ".join(a), " ".join(b)).ratio() >= 0.92


def confirmed_document(group, credit):
    documents = {digits(p.beneficiary_document) for p in group.payments}
    document, kind = validate_document(credit.cedent_document)
    return kind is not None and documents == {document}


def infer_block_references(groups, existing, failure_key, lawyer_rules):
    """Bare lawyer names require one principal in the same physical PFMI block."""
    partitions = defaultdict(list)
    for group in groups:
        p = group.first
        partitions[p.file_id, p.source_sheet, p.block_id, failure_key(p.failure)].append(group)
    references = {}
    for items in partitions.values():
        if len(items) < 2 or any(g.group_id in existing for g in items):
            continue
        # JOSE VALDIR is the explicit bare-name case supplied for this workflow.
        known = {normalize_name("JOSE VALDIR")}
        for g in items:
            known.update(lawyer_rules.get(failure_key(g.first.failure), ()))
        satellites = [g for g in items if normalize_name(g.first.cedent_name) in known]
        principals = [g for g in items if g not in satellites]
        if len(principals) == 1 and satellites and len(satellites) == len(items) - 1:
            references.update({g.group_id: principals[0].group_id for g in satellites})
    return references


def document_composition_candidate(
    principal, components, analytic, calculation, total, failure_key,
):
    """Resolve a known document within its failure before attempting a value-based search."""
    if principal.first.modality == CESSAO or total is None:
        return None, None
    if len({(g.first.modality, g.first.commission_text) for g in components}) != 1:
        return None, None
    identities = [
        credit for credit in analytic
        if failure_key(credit.campaign) == failure_key(principal.first.failure)
        and confirmed_document(principal, credit)
    ]
    if not identities:
        return None, None
    limit = tolerance_cents()
    candidates = []
    for credit in identities:
        expected = calculation(credit, principal)[3]
        if expected is not None and expected >= 0 and reconcile_credit(
            "", "", "", total, expected, limit_cents=limit,
        ).status != "CRITICAL_ERROR":
            candidates.append(credit)
    if len(candidates) > 1:
        return None, "VALOR_AMBIGUO"
    if not candidates:
        return None, "DOCUMENTO_VALOR_DIVERGENTE"
    credit = candidates[0]
    if not credit.source_sheet or credit.source_row < 2:
        return None, "LOCALIZACAO_ANALITICO_AUSENTE"
    return credit, "DOCUMENTO_FALENCIA_VALOR_UNICOS"


def smart_candidate(principal, components, analytic, calculation, total):
    """Require one financial candidate before checking independent identity evidence."""
    if principal.first.modality == CESSAO or total is None:
        return None, "MODALIDADE_OU_TOTAL_NAO_SUPORTADO"
    if len({(g.first.modality, g.first.commission_text) for g in components}) != 1:
        return None, "PARAMETROS_DOS_COMPONENTES_DIVERGENTES"

    candidates = []
    limit = tolerance_cents()
    for credit in analytic:
        expected = calculation(credit, principal)[3]
        if expected is None or expected < 0:
            continue
        result = reconcile_credit("", "", "", total, expected, limit_cents=limit)
        if result.status != "CRITICAL_ERROR":
            candidates.append(credit)
    if not candidates:
        return None, "VALOR_NAO_LOCALIZADO"
    if len(candidates) > 1:
        return None, "VALOR_AMBIGUO"

    credit = candidates[0]
    doc_match = confirmed_document(principal, credit)
    if not doc_match and not strong_name_match(principal.first.cedent_name, credit.cedent_name):
        return None, "IDENTIDADE_INSUFICIENTE"
    if validate_document(credit.cedent_document)[1] is None:
        return None, "DOCUMENTO_ANALITICO_INVALIDO"
    if not credit.source_sheet or credit.source_row < 2:
        return None, "LOCALIZACAO_ANALITICO_AUSENTE"
    return credit, "DOCUMENTO_EXATO" if doc_match else "NOME_FORTE"


def audit_fill(values, field, value, rule, source):
    records = json.loads(values.get("PREENCHIMENTOS_AUTOMATICOS") or "[]")
    records.append({
        "campo": field, "valor_anterior": values.get(field, ""), "valor_sugerido": value,
        "regra": rule, "fonte": source, "exige_aprovacao_humana": False,
    })
    values[field] = value
    values["PREENCHIMENTOS_AUTOMATICOS"] = json.dumps(records, ensure_ascii=False, default=str)
    values["ALERTAS"] = str(values.get("ALERTAS") or "") + f"; {rule}: {field}={value}"


def system_approval_valid(values, originals, field):
    """An editable marker cannot manufacture system provenance or approve new edits."""
    if not has_system_approval(originals, field):
        return False
    if field == "APROVADO":
        fields = ["NOME_CEDENTE", "DOC_CEDENTE"]
    else:
        fields = ["COMPOSICAO_SELECAO_MANUAL_ABA", "COMPOSICAO_SELECAO_MANUAL_LINHA"]
        if field == "COMPOSICAO_APROVADA":
            fields += ["VL_NOMINAL", "VL_PRESENTE"]
    # The canonical representation is shared with the protected workbook snapshot.
    from .workbook import _canonical

    return all(_canonical(values.get(k)) == _canonical(originals.get(k)) for k in fields)


def has_system_approval(originals, field):
    if originals.get(field) == SYSTEM_APPROVAL:
        return True
    if originals.get(field) != "SIM":
        return False
    try:
        records = json.loads(originals.get("PREENCHIMENTOS_AUTOMATICOS") or "[]")
        return any(r.get("campo") == field and r.get("valor_sugerido") == "SIM"
                   and r.get("regra") == "AUTOMATED_MATCH_APPLIED" for r in records)
    except (ValueError, TypeError, AttributeError):
        return False


def suggested_credit(text, group, analytic, calculation, failure_key):
    """A parsed location is a lead, never sufficient evidence for approval."""
    locations = {(sheet.strip(), int(row)) for sheet, row in re.findall(
        r"Aba:\s*([^\r\n,]+?)\s*,\s*Linha:\s*(\d+)\b", text, flags=re.IGNORECASE,
    )}
    if not locations or group.total is None:
        return None

    def failure_tokens(value):
        tokens = normalize_name(failure_key(value)).split()
        while len(tokens) > 2 and tokens[-1] in {"SA", "LTDA"}:
            tokens.pop()
        return tokens

    compatible = []
    for credit in analytic:
        if failure_tokens(credit.campaign) != failure_tokens(group.first.failure):
            continue
        same_doc = confirmed_document(group, credit)
        same_name = normalize_name(credit.cedent_name) == normalize_name(group.first.cedent_name)
        if not (same_doc or (group.first.modality == CESSAO and same_name)):
            continue
        expected = calculation(credit, group)[3]
        if expected is None or expected < 0 or reconcile_credit(
            "", "", "", group.total, expected,
        ).status == "CRITICAL_ERROR":
            continue
        compatible.append(credit)
    if len(compatible) != 1:
        return None
    credit = compatible[0]
    location = (credit.source_sheet, credit.source_row)
    if location not in locations or credit.source_row < 2 or not credit.source_sheet:
        return None
    if sum((c.source_sheet, c.source_row) == location for c in analytic) != 1:
        return None
    if validate_document(credit.cedent_document)[1] is None:
        return None
    return credit
