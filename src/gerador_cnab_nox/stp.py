"""Offline STP evidence rules. Monetary coincidence alone never proves identity."""

from __future__ import annotations

import json
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
    if originals.get(field) != SYSTEM_APPROVAL:
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
