from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict
from dataclasses import asdict, dataclass
from decimal import ROUND_HALF_UP, Decimal, localcontext

from .errors import InputFileError
from .failure_aliases import resolver
from .models import (
    CESSAO,
    CONTROL_FIELDS,
    FINAL_FIELDS,
    AnalyticCredit,
    DueRecord,
    PfmiData,
    PfmiPayment,
    PreparedBatch,
    PreparedRow,
)
from .normalize import (
    calculate_commission,
    commission_rate,
    digits,
    mask_document,
    names_equivalent_under_confirmed_document,
    normalize_failure,
    normalize_name,
    parse_date,
    technical_name,
    validate_document,
)
from .reconciliation import reconcile_credit
from .stp import (
    SYSTEM_APPROVAL,
    audit_fill,
    confirmed_document,
    infer_block_references,
    smart_candidate,
)

CONEXCRED_NAME = "CONEXCRED INTERMEDIACAO"
# Pendências que já dizem exatamente o que o operador precisa fazer - quando
# alguma delas está presente, o SELECAO_MANUAL_NECESSARIA genérico é
# redundante (ver dedup em prepare_batch).
_SPECIFIC_ACTION_PENDENCIES = frozenset(
    {"DIVERGENCIA_VALOR", "MULTIPLOS_CREDITOS_PRINCIPAIS", "DOCUMENTOS_CANDIDATOS_CONFLITANTES"}
)


def _is_conexcred_beneficiary(value):
    # Exact "CONEXCRED" or a name that starts with "CONEXCRED " (word boundary).
    # A name that merely contains "conexcred" mid-word must never match.
    name = normalize_name(value)
    return name == "CONEXCRED" or name.startswith("CONEXCRED ")


@dataclass
class _Group:
    group_id: str
    order: int
    payments: list[PfmiPayment]

    @property
    def first(self):
        return self.payments[0]

    @property
    def total(self):
        if any(p.beneficiary_value is None for p in self.payments):
            return None
        return sum((p.beneficiary_value for p in self.payments), Decimal("0.00"))


def _group_payments(payments, failure_key=normalize_failure):
    grouped = {}
    for p in sorted(payments, key=lambda p: p.file_order):
        key = (
            p.file_id,
            p.source_sheet,
            p.block_id,
            normalize_name(p.cedent_name),
            failure_key(p.failure),
        )
        grouped.setdefault(key, []).append(p)
    return [
        _Group(f"{items[0].file_id}-{items[0].block_id}-G{n:04d}", n, items)
        for n, items in enumerate(grouped.values(), 1)
    ]


def lawyer_name(value, failure, rules):
    # Only a prefix matching an identity confirmed in the local rules is usable.
    # Content inside parentheses never selects a credit or supplies a document.
    # No rule for this failure -> "" -> falls through to manual review.
    prefix = str(value or "").split("(", 1)[0].strip()
    key = normalize_name(prefix)
    identities = rules.get(normalize_failure(failure), ())
    if any(key == identity or key.startswith(identity + " ") for identity in identities):
        return prefix
    return ""


def is_lawyer_case(failure, rules):
    return normalize_failure(failure) in rules


def lawyer_identity(value, failure, rules):
    name = normalize_name(lawyer_name(value, failure, rules))
    for identity in rules.get(normalize_failure(failure), ()):
        if name.startswith(identity):
            return identity
    return ""


def _calculation(credit, group):
    p = group.first
    try:
        rate = commission_rate(p.commission_text, p.modality)
    except ValueError as exc:
        return None, None, None, None, str(exc)
    try:
        base, commission, expected = calculate_commission(
            credit.nominal_value, credit.acquisition_value, rate
        )
        return rate, base, commission, expected, ""
    except ValueError as exc:
        return rate, None, None, None, str(exc)


def _related_block_suggestions(groups, calculations, analytic, failure_key):
    """Exploratory only: parentheses match + exact aggregate value. Never selects."""
    suggestions = []
    seen_pairs = set()
    for outer in groups:
        for inner in groups:
            if outer.group_id == inner.group_id:
                continue
            if outer.first.file_id != inner.first.file_id:
                continue
            if failure_key(outer.first.failure) != failure_key(inner.first.failure):
                continue
            pair_key = frozenset((outer.group_id, inner.group_id))
            if pair_key in seen_pairs:
                continue
            match = re.search(r"\(([^()]+)\)", inner.first.cedent_name)
            if not match or normalize_name(match.group(1)) != normalize_name(
                outer.first.cedent_name
            ):
                continue
            if outer.total is None or inner.total is None:
                continue
            total = outer.total + inner.total
            found = [
                c
                for c in analytic
                if failure_key(c.campaign) == failure_key(outer.first.failure)
                and calculations.get(
                    (outer.group_id, c.credit_id), (None, None, None, None, "")
                )[3]
                == total
            ]
            if not found:
                continue
            seen_pairs.add(pair_key)
            suggestions.append(
                {
                    "grupo_a": outer.group_id,
                    "credor_a": outer.first.cedent_name,
                    "valor_a": outer.total,
                    "grupo_b": inner.group_id,
                    "credor_b": inner.first.cedent_name,
                    "valor_b": inner.total,
                    "total": total,
                    "creditos_encontrados": [
                        {"referencia": c.reference, "valor": total} for c in found
                    ],
                }
            )
    return suggestions


def _composition_manual_candidates(principal, analytic, total_pfmi):
    """Exploratory only, for human review in the Excel: every Analítico
    credit whose expected acquisition (principal's own commission rate,
    Decimal-exact, no rounding) equals the composition's PFMI total exactly.

    Never restricted to a shared campaign/batch field (some real Analítico
    files reuse the same "Prospect" text across hundreds of unrelated
    credits, which is never a safe identifier) and never selects anything by
    itself: the operator still confirms one specific record (aba + linha,
    see COMPOSICAO_SELECAO_MANUAL_ABA/LINHA in models.py) after checking the
    physical instrument.
    """
    try:
        rate = commission_rate(principal.first.commission_text, principal.first.modality)
    except ValueError:
        return []
    candidates = []
    for c in analytic:
        try:
            _, _, expected = calculate_commission(c.nominal_value, c.acquisition_value, rate)
        except ValueError:
            continue
        if expected is not None and expected == total_pfmi:
            candidates.append(
                {
                    "aba": c.source_sheet,
                    "linha": c.source_row,
                    "nome": c.cedent_name,
                    "documento_mascarado": mask_document(c.cedent_document),
                    "nominal": c.nominal_value,
                    "aquisicao": c.acquisition_value,
                }
            )
    return candidates


def _composition_candidates(groups, primary, match_rule, analytic, failure_key, lawyer_rules):
    """Deterministic, structural-only proposal that N PFMI groups (payments)
    together represent ONE credit of the Analitico.

    Never selects, never confirms identity, never fills INCLUIR_CNAB: human
    approval is exclusively the COMPOSICAO_APROVADA field. Star topology only
    (every satellite references the same principal directly; no chains), all
    components share the PFMI file and the normalized failure, and the credit
    match reuses the same deterministic name+failure rule already used for
    ordinary eligibility ("primary") - never a global search over values.
    """
    by_id = {g.group_id: g for g in groups}
    reference_of: dict[str, str] = {}
    for g in groups:
        match = re.search(r"\(([^()]+)\)", g.first.cedent_name)
        if not match:
            continue
        target = normalize_name(match.group(1))
        hits = [
            other
            for other in groups
            if other.group_id != g.group_id
            and other.first.file_id == g.first.file_id
            and failure_key(other.first.failure) == failure_key(g.first.failure)
            and normalize_name(other.first.cedent_name) == target
        ]
        if len(hits) == 1:
            reference_of[g.group_id] = hits[0].group_id
        elif len(hits) > 1:
            reference_of[g.group_id] = "AMBIGUA"

    inferred = infer_block_references(groups, reference_of, failure_key, lawyer_rules)
    reference_of.update(inferred)
    satellites_of: dict[str, list[str]] = defaultdict(list)
    for satellite_id, principal_id in reference_of.items():
        if principal_id != "AMBIGUA":
            satellites_of[principal_id].append(satellite_id)

    compositions: list[dict] = []
    used_groups: set[str] = set()
    for principal_id in sorted(satellites_of):
        if principal_id in reference_of:
            # A "principal" that itself references someone else is not a root:
            # reject deterministically instead of guessing a reference chain.
            continue
        principal = by_id[principal_id]
        satellite_ids = sorted(satellites_of[principal_id])
        component_ids = [principal_id, *satellite_ids]
        if used_groups.intersection(component_ids):
            continue  # each group belongs to at most one composition
        components = [by_id[cid] for cid in component_ids]
        composition_id = hashlib.sha256(
            f"COMPOSICAO|{'|'.join(component_ids)}".encode()
        ).hexdigest()[:20].upper()
        participants_text = "; ".join(g.first.cedent_name for g in components)
        base = {
            "composicao_id": composition_id,
            "component_ids": component_ids,
            "participantes": participants_text,
            "regra": "REFERENCIA_PARENTESE_MESMO_ARQUIVO_FALENCIA",
            "total_pfmi": None,
            "total_analitico": None,
            "diferenca": None,
            "nominal_analitico": None,
            "credito_referencia": "",
            "credit": None,
            "candidatos_localizacao_manual": [],
            "auto_approved": False,
        }
        if any(cid in inferred for cid in component_ids):
            base["regra"] = "TITULAR_UNICO_ADVOGADOS_MESMO_BLOCO"

        def _reject(state, reason, extra=None, component_ids=component_ids, base=base):
            used_groups.update(component_ids)
            compositions.append({**base, **(extra or {}), "estado": state, "motivo": reason})

        ambiguous = any(
            reference_of.get(g.group_id) == "AMBIGUA"
            and re.search(r"\(([^()]+)\)", g.first.cedent_name)
            and normalize_name(re.search(r"\(([^()]+)\)", g.first.cedent_name).group(1))
            == normalize_name(principal.first.cedent_name)
            for g in groups
        )
        if ambiguous:
            _reject(
                "BLOQUEADA_REFERENCIA_AMBIGUA",
                "Referência entre parênteses corresponde a mais de um grupo do mesmo arquivo",
            )
            continue
        if any(g.total is None or any(p.issues for p in g.payments) for g in components):
            _reject(
                "BLOQUEADA_PENDENCIA_ORIGEM",
                "Um ou mais componentes têm pendência de origem no PFMI",
            )
            continue
        invalid_doc_group = next(
            (
                g
                for g in components
                if any(validate_document(p.beneficiary_document)[1] is None for p in g.payments)
            ),
            None,
        )
        if invalid_doc_group is not None:
            _reject(
                "BLOQUEADA_DOCUMENTO_INVALIDO",
                f"Documento de beneficiário inválido em {invalid_doc_group.group_id}",
            )
            continue
        # O matching comum (primary) já tenta nome e, quando não bate, um
        # único documento válido idêntico na mesma falência (matching.py,
        # loop principal de prepare_batch) - a composição só reaproveita o
        # resultado, nunca duplica ou amplia essa busca.
        main_credits = list(primary.get(principal_id, []))
        principal_match_rule = match_rule.get(principal_id, "NOME")
        total_pfmi = sum((g.total for g in components), Decimal("0.00"))
        smart_credit, smart_rule = smart_candidate(
            principal, components, analytic, _calculation, total_pfmi,
        )
        base["stp_rule"] = smart_rule
        if not main_credits and smart_credit is not None:
            main_credits = [smart_credit]
            principal_match_rule = "SMART_MATCH"
        if len(main_credits) == 0:
            total_pfmi_sem_principal = sum((g.total for g in components), Decimal("0.00"))
            candidatos = _composition_manual_candidates(
                principal, analytic, total_pfmi_sem_principal
            )
            _reject(
                "BLOQUEADA_CREDITO_NAO_LOCALIZADO",
                "Nenhum crédito do Analítico corresponde ao principal (nome ou documento)",
                extra={"candidatos_localizacao_manual": candidatos},
            )
            continue
        if len(main_credits) > 1:
            _reject(
                "BLOQUEADA_MULTIPLOS_CREDITOS",
                "Mais de um crédito do Analítico corresponde ao principal",
            )
            continue
        credit = main_credits[0]
        base["auto_approved"] = (
            smart_credit is not None and smart_credit.credit_id == credit.credit_id
            and credit.nominal_value is not None and credit.nominal_value > 0
            and credit.signature_date is not None
            and not any(
                credit.credit_id in {c.credit_id for c in primary.get(other.group_id, [])}
                for other in groups if other.group_id not in component_ids
            )
        )
        # A credit reused by another group only reaches here through an exact
        # duplicate of the principal's own name, which the ambiguity check
        # above always rejects first; the pre-existing multi-group candidate
        # protection (CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS) still applies
        # unmodified to any row, composition or not.
        try:
            rate = commission_rate(principal.first.commission_text, principal.first.modality)
            _, _, expected = calculate_commission(
                credit.nominal_value, credit.acquisition_value, rate
            )
        except ValueError as exc:
            _reject("BLOQUEADA_PENDENCIA_ORIGEM", str(exc))
            continue
        if expected is None:
            _reject(
                "BLOQUEADA_PENDENCIA_ORIGEM",
                "Não foi possível calcular a aquisição esperada do crédito",
            )
            continue
        total_pfmi = sum((g.total for g in components), Decimal("0.00"))
        diferenca = total_pfmi - expected
        used_groups.update(component_ids)
        doc_note = (
            " Principal localizado por documento (CPF/CNPJ), não por nome: "
            "corrija NOME_CEDENTE e marque APROVADO=SIM por divergência de nome."
            if principal_match_rule == "DOCUMENTO"
            else ""
        )
        compositions.append(
            {
                **base,
                "estado": "PROPOSTA" if diferenca == 0 else "BLOQUEADA_DIVERGENCIA_VALOR",
                "motivo": (
                    "Soma exata do PFMI confere com a aquisição esperada do crédito" + doc_note
                    if diferenca == 0
                    else (
                        f"Diferença de {diferenca:+.2f} entre soma do PFMI e aquisição esperada"
                        + doc_note
                    )
                ),
                "total_pfmi": total_pfmi,
                "total_analitico": expected,
                "diferenca": diferenca,
                "nominal_analitico": credit.nominal_value,
                "credito_referencia": credit.reference,
                # O principal já está identificado sem ambiguidade (um único
                # credito, nome/documento conferidos acima) mesmo quando só o
                # VALOR diverge - mantém o crédito disponível para que a Lu
                # veja o nome/documento/nominal sugeridos e, se aprovar
                # explicitamente a divergência de valor (validation.py), gere
                # com o valor do PFMI. Nunca guarda o crédito quando a causa
                # do bloqueio é outra (ambíguo, documento inválido, etc.).
                "credit": credit,
            }
        )
    return compositions


def prepare_batch(
    pfmi: PfmiData,
    analytic: list[AnalyticCredit],
    due_base: list[DueRecord],
    *,
    liquidation_date,
    first_sequence,
    failure_aliases=None,
    lawyer_rules=None,
    debtor_fallbacks=None,
) -> PreparedBatch:
    lawyer_rules = lawyer_rules or {}
    failure_key = resolver(failure_aliases or {})
    resolved_fallbacks = {}
    for key, entry in (debtor_fallbacks or {}).items():
        canonical = failure_key(key)
        if canonical in resolved_fallbacks and resolved_fallbacks[canonical] != entry:
            raise InputFileError("Equivalências apontam para fallbacks de sacado conflitantes")
        resolved_fallbacks[canonical] = entry
    debtor_fallbacks = resolved_fallbacks
    groups = _group_payments(pfmi.payments, failure_key)
    due_index = _index_due_base(due_base, failure_key)
    calculations, primary, candidates, eligible, match_rule = {}, {}, {}, {}, {}
    for g in groups:
        same_failure = [
            c
            for c in analytic
            if failure_key(c.campaign) == failure_key(g.first.failure)
        ]
        cedent_name_norm = normalize_name(g.first.cedent_name)
        main = [c for c in same_failure if normalize_name(c.cedent_name) == cedent_name_norm]
        rule = "NOME"
        if not main:
            # Nome não bate: um único crédito com o MESMO documento válido
            # (CPF/CNPJ) do beneficiário, na MESMA falência, é aceito como
            # alternativa - nunca por similaridade, nunca atravessando
            # falências. Nunca torna a linha elegível automaticamente: exige
            # sempre revisão humana (INCLUIR_CNAB manual) e APROVADO=SIM pela
            # divergência de nome, como qualquer outra DIVERGENCIA_NOME.
            beneficiary_doc = digits(g.first.beneficiary_document)
            if beneficiary_doc and validate_document(beneficiary_doc)[1] is not None:
                doc_matches = [
                    c
                    for c in same_failure
                    if digits(c.cedent_document) == beneficiary_doc
                    and validate_document(c.cedent_document)[1] is not None
                ]
                # Ambíguo por documento (2+ créditos): nunca populariza como
                # candidato automático - fica como crédito não localizado,
                # igual a hoje, em vez de arriscar linhas extras imprevistas.
                if len(doc_matches) == 1:
                    main = doc_matches
                    rule = "DOCUMENTO"
        match_rule[g.group_id] = rule
        primary[g.group_id] = main
        candidates[g.group_id] = list(main)
        for c in same_failure:
            calculations[g.group_id, c.credit_id] = _calculation(c, g)
        docs = [validate_document(c.cedent_document) for c in main]
        expected = [calculations[g.group_id, c.credit_id][3] for c in main]
        eligible[g.group_id] = (
            rule in {"NOME", "DOCUMENTO"}
            and len(main) == 1
            and all(t is not None for _, t in docs)
            and len({d for d, _ in docs}) == 1
            and all(e is not None for e in expected)
            and sum((e for e in expected if e is not None), Decimal(0)) == g.total
            and not any(p.issues for p in g.payments)
        )

    # Purely observational: computed from copies, wrapped so it can never affect
    # groups/candidates/eligible/status below, even on an unexpected error.
    try:
        related_suggestions = _related_block_suggestions(
            list(groups), dict(calculations), list(analytic), failure_key
        )
    except Exception:
        related_suggestions = []

    compositions = _composition_candidates(
        list(groups), dict(primary), dict(match_rule), list(analytic), failure_key, lawyer_rules,
    )
    # One analytic credit cannot be automatically assigned to two compositions.
    auto_uses = defaultdict(list)
    for composition in compositions:
        if composition.get("auto_approved"):
            auto_uses[composition["credit"].credit_id].append(composition)
    for uses_for_credit in auto_uses.values():
        if len(uses_for_credit) > 1:
            for composition in uses_for_credit:
                composition["auto_approved"] = False
                composition["stp_rule"] = "CREDITO_REUTILIZADO"
    composition_by_group: dict[str, dict] = {}
    for composition in compositions:
        composition["nominais_sugeridos"] = _suggest_composition_nominals(
            composition, {g.group_id: g for g in groups},
        )
        if composition.get("auto_approved") and len(composition["nominais_sugeridos"]) != len(
            composition["component_ids"]
        ):
            composition["auto_approved"] = False
            composition["stp_rule"] = "NOMINAIS_INDISPONIVEIS"
        for group_id in composition["component_ids"]:
            composition_by_group[group_id] = composition

    expanded = set()
    while True:
        uses = defaultdict(set)
        for gid, credits in candidates.items():
            for c in credits:
                uses[c.credit_id].add(gid)
        pending = [
            g
            for g in groups
            if g.group_id not in expanded
            and (
                not eligible[g.group_id]
                or any(len(uses[c.credit_id]) > 1 for c in primary[g.group_id])
            )
        ]
        if not pending:
            break
        for g in pending:
            expanded.add(g.group_id)
            known = {c.credit_id for c in candidates[g.group_id]}
            for c in analytic:
                calc = calculations.get((g.group_id, c.credit_id))
                if c.credit_id not in known and calc and calc[3] is not None and calc[3] == g.total:
                    candidates[g.group_id].append(c)
    prepared = []
    for g in groups:
        main_ids = {c.credit_id for c in primary[g.group_id]}
        shared = any(len(uses[c.credit_id]) > 1 for c in candidates[g.group_id])
        auto = eligible[g.group_id] and not shared
        status, due, due_issues = _resolve_due(due_index, g.first.failure, failure_key)
        fallback = (debtor_fallbacks or {}).get(failure_key(g.first.failure)) if status else None
        if fallback:
            due = DueRecord(0, fallback["name"], fallback["document"], liquidation_date)
            status, due, due_issues = _resolve_due(
                {failure_key(g.first.failure): [due]}, g.first.failure, failure_key,
            )
        credits = sorted(candidates[g.group_id], key=lambda c: c.source_row)
        # Um componente de composição já PROPOSTA (soma exata, um único
        # principal encontrado) não é um crédito individual: o matching comum
        # (nome/documento/valor por grupo isolado) não se aplica a ele - quem
        # decide identidade e valor é a composição inteira (ver
        # _composition_candidates), nunca o total isolado deste grupo. Sem
        # isso, o principal aparecia com DIVERGENCIA_VALOR (seu total sozinho
        # nunca fecha com o crédito) e o satélite exigia seleção manual
        # individual, mesmo com a composição já fechando exata.
        composition_for_group = composition_by_group.get(g.group_id)
        if composition_for_group and composition_for_group.get("auto_approved"):
            credits = (
                [composition_for_group["credit"]]
                if g.group_id == composition_for_group["component_ids"][0] else []
            )
            for credit in credits:
                calculations[g.group_id, credit.credit_id] = _calculation(credit, g)
        is_composition_resolved = (
            bool(composition_for_group) and (
                composition_for_group["estado"] == "PROPOSTA"
                or composition_for_group.get("auto_approved")
            )
        )
        # Uma composição bloqueada (qualquer estado BLOQUEADA_* exceto
        # divergência de valor, que já tem sua própria pendência específica)
        # é uma causa raiz única para os componentes sem crédito próprio:
        # apontar SELECAO_MANUAL_NECESSARIA e CREDITO_NAO_LOCALIZADO lado a
        # lado seria a mesma causa contada duas vezes.
        is_composition_blocked = bool(composition_for_group) and composition_for_group[
            "estado"
        ] not in ("PROPOSTA", "BLOQUEADA_DIVERGENCIA_VALOR")
        for c in credits or [None]:
            issues = list(due_issues)
            composition_blocked_root = is_composition_blocked and c is None
            if composition_blocked_root:
                issues.append("COMPOSICAO_BLOQUEADA_SELECAO_MANUAL")
            elif not auto and not is_composition_resolved:
                issues.append("SELECAO_MANUAL_NECESSARIA")
            if status:
                issues.append(status)
            reason = "PRINCIPAL_NOME_FALENCIA_CONJUNTO_INTEGRAL"
            if c is None:
                if is_composition_resolved:
                    # Satélite de composição já PROPOSTA: nunca teve (nem
                    # precisa de) crédito próprio - o rótulo não pode soar
                    # como uma pendência real, já que o crédito compartilhado
                    # já está resolvido (ver analytic_credit em _make_row).
                    reason = "COMPOSICAO_SATELITE_CREDITO_COMPARTILHADO"
                elif composition_blocked_root:
                    reason = "COMPOSICAO_BLOQUEADA_SELECAO_MANUAL"
                else:
                    issues.append("CREDITO_NAO_LOCALIZADO")
                    reason = "CREDITO_NAO_LOCALIZADO"
            else:
                if match_rule.get(g.group_id) == "DOCUMENTO" and c.credit_id in main_ids:
                    issues.append("CREDOR_LOCALIZADO_POR_DOCUMENTO")
                calc = calculations[g.group_id, c.credit_id]
                if calc[4]:
                    issues.append(calc[4])
                if c.credit_id not in main_ids:
                    reason = "ALTERNATIVA_VALOR_EXATO_ESCOLHA_MANUAL"
                    issues.append("DIVERGENCIA_NOME")
                elif not eligible[g.group_id] and not is_composition_resolved:
                    expected = [
                        calculations[g.group_id, item.credit_id][3] for item in primary[g.group_id]
                    ]
                    if len(primary[g.group_id]) > 1:
                        issues.append("MULTIPLOS_CREDITOS_PRINCIPAIS")
                    if all(e is not None for e in expected) and sum(expected) != g.total:
                        issues.append("DIVERGENCIA_VALOR")
                    if len({digits(item.cedent_document) for item in primary[g.group_id]}) > 1:
                        issues.append("DOCUMENTOS_CANDIDATOS_CONFLITANTES")
                if len(uses[c.credit_id]) > 1:
                    issues.append("CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS")
                for value, label in (
                    (c.nominal_value, "VL_NOMINAL"),
                    (c.acquisition_value, "VL_PRESENTE"),
                ):
                    if value is None:
                        issues.append(label + "_INVALIDO")
                    elif value < 0:
                        issues.append(label + "_NEGATIVO")
                if g.first.modality != CESSAO:
                    # In CESSAO, the final cedente document is Conexcred's CNPJ,
                    # validated separately below; the original credit's document
                    # (often an internal code, not a CPF/CNPJ) must not block it.
                    if validate_document(c.cedent_document)[1] is None:
                        issues.append("DOC_CEDENTE_INVALIDO")
                    if c.signature_date is None:
                        issues.append("DT_EMISSAO_INVALIDA")
                    elif liquidation_date and c.signature_date > liquidation_date:
                        issues.append("DT_EMISSAO_POSTERIOR_LIQUIDACAO")
            # SELECAO_MANUAL_NECESSARIA só diz "escolha manual" em termos
            # genéricos; quando a mesma linha já carrega uma pendência
            # específica que já diz exatamente o que fazer (ex.:
            # DIVERGENCIA_VALOR, MULTIPLOS_CREDITOS_PRINCIPAIS), prefixá-la
            # de novo não soma informação, só duplica a mesma causa.
            if "SELECAO_MANUAL_NECESSARIA" in issues and _SPECIFIC_ACTION_PENDENCIES.intersection(
                issues
            ):
                issues = [i for i in issues if i != "SELECAO_MANUAL_NECESSARIA"]
            row = _make_row(
                    g,
                    c,
                    due,
                    due_index,
                    issues,
                    reason,
                    "SIM" if auto and c and c.credit_id in main_ids else "NAO",
                    [pay for pay in pfmi.payments if pay.file_id == g.first.file_id],
                    failure_key,
                    lawyer_rules,
                    liquidation_date,
                    composition_by_group.get(g.group_id),
            )
            if fallback:
                for field in ("DOC_SACADO", "NOME_SACADO", "DT_VENCIMENTO"):
                    audit_fill(row.values, field, row.values[field],
                               "FALLBACK_SACADO_DATA_LIQUIDACAO", fallback["source"])
                row.originals.update(row.values)
            prepared.append(row)
    # Each GROUP reserves one sequence number (order of PFMIs, then order of
    # groups), shared by every alternative/no-candidate row of that group:
    # only one of them is ever meant to be selected. A group whose real
    # credits are multiple complementary principals (same document, sum
    # exact) reserves one extra number per extra principal, since those are
    # selected together as distinct CNAB titles. Approving a different
    # candidate later only edits INCLUIR_CNAB/APROVADO: numbers already fit.
    if isinstance(first_sequence, int):
        sequence, current_group, group_first, principals_seen = (
            first_sequence,
            None,
            None,
            0,
        )
        for row in prepared:
            group_id = row.values["ID_GRUPO"]
            if group_id != current_group:
                if current_group is not None:
                    sequence += 1
                current_group, group_first, principals_seen = group_id, sequence, 0
            is_principal = (
                row.values["MOTIVO_CORRESPONDENCIA"] == "PRINCIPAL_NOME_FALENCIA_CONJUNTO_INTEGRAL"
            )
            if is_principal:
                if principals_seen:
                    sequence += 1
                assigned, principals_seen = sequence, principals_seen + 1
            else:
                assigned = group_first
            for field in ("SEU_NUMERO", "NU_DOCUMENTO"):
                row.values[field] = row.originals[field] = assigned
    return PreparedBatch(
        prepared,
        pfmi.unknown_rows,
        liquidation_date,
        first_sequence,
        warnings=list(dict.fromkeys(e.get("alerta") for e in pfmi.evidence if e.get("alerta"))),
        payment_count=len(pfmi.payments),
        group_count=len(groups),
        evidence=pfmi.evidence,
        sources=pfmi.sources,
        payments_snapshot=[asdict(p) for p in pfmi.payments],
        related_suggestions=related_suggestions,
        compositions=compositions,
    )


def _index_due_base(records, failure_key=normalize_failure):
    index = defaultdict(list)
    for r in records:
        index[failure_key(r.debtor_name)].append(r)
    return index


def _resolve_due(index, failure, failure_key=normalize_failure):
    records = index.get(failure_key(failure), [])
    if not records:
        return "PREENCHER_MANUALMENTE", None, ["SACADO_NAO_LOCALIZADO"]
    distinct = {}
    for r in records:
        key = (normalize_failure(r.debtor_name), digits(r.debtor_document), r.due_date)
        distinct.setdefault(key, r)
    if len(distinct) > 1:
        return "CONFLITO_BASE_VENCIMENTOS", None, ["CONFLITO_BASE_VENCIMENTOS"]
    r = next(iter(distinct.values()))
    issues = []
    if validate_document(r.debtor_document)[1] is None:
        issues.append("DOC_SACADO_INVALIDO")
    if r.due_date is None:
        issues.append("DATA_VENCIMENTO_INVALIDA")
    return "PREENCHER_MANUALMENTE" if issues else "", r, issues


def _suggest_composition_nominals(composition, groups_by_id):
    """Proportional suggestion only; never selects or approves a component."""
    total_raw = composition.get("total_analitico")
    nominal_raw = composition.get("nominal_analitico")
    if total_raw is None or nominal_raw is None:
        return {}
    total, nominal = Decimal(str(total_raw)), Decimal(str(nominal_raw))
    if not total.is_finite() or not nominal.is_finite() or total <= 0 or nominal <= 0:
        return {}
    component_ids = composition["component_ids"]
    amounts = [groups_by_id[gid].total for gid in component_ids]
    if not component_ids or any(
        amount is None or not amount.is_finite() or amount < 0 for amount in amounts
    ):
        return {}
    with localcontext() as context:
        context.prec = 80
        cent = Decimal("0.01")
        # Source nominal must already represent a monetary value in cents.
        if nominal != nominal.quantize(cent, rounding=ROUND_HALF_UP):
            return {}
        values = [
            (nominal * (Decimal(str(amount)) / total)).quantize(cent, rounding=ROUND_HALF_UP)
            for amount in amounts
        ]
        # component_ids[0] is the detected star's principal, regardless of row order.
        values[0] += nominal - sum(values, Decimal("0.00"))
        if any(value < 0 for value in values):
            return {}
        assert sum(values, Decimal("0.00")) == nominal
    return dict(zip(component_ids, values, strict=True))


def _make_row(g, c, due, due_index, issues, reason, selected, file_payments,
              failure_key=normalize_failure, lawyer_rules=None, liquidation_date=None,
              composition=None):
    lawyer_rules = lawyer_rules or {}
    p = g.first
    credit_id = c.credit_id if c else ""
    line_id = hashlib.sha256(f"{g.group_id}|{credit_id}".encode()).hexdigest()[:20].upper()
    preenchimentos_automaticos: list[dict] = []
    is_composition_principal = bool(composition) and composition["component_ids"][0] == g.group_id
    is_composition_satellite = bool(composition) and not is_composition_principal
    # A composition satellite never has a credit of its own for CNAB linkage
    # purposes: even when the ordinary per-group matching (primary[]) happens
    # to find an independent candidate for it (same name/document by
    # coincidence), that candidate is a DIFFERENT real credit and using it
    # here would silently swap the composition's shared credit for an
    # unrelated one - explicitly rejected by design. Once the composition has
    # a resolved shared credit (PROPOSTA/BLOQUEADA_DIVERGENCIA_VALOR), every
    # component (principal and satellites) traces to that SAME credit. A
    # principal without a resolved composition credit yet (e.g. multiple
    # ambiguous candidates) still falls back to its own ordinary match `c`,
    # exactly like before, so the existing "pick one of the candidate rows"
    # disambiguation keeps working.
    if composition:
        analytic_credit = composition["credit"] or (c if is_composition_principal else None)
    else:
        analytic_credit = c
    # Sem nenhum crédito resolvido (nem individual, nem compartilhado por
    # composição), DOC_CEDENTE/NOME_CEDENTE ficam vazios só porque não há de
    # onde tirá-los - um satélite não conta aqui: ele nunca depende de
    # analytic_credit, sua própria pendência de documento/nome (quando o
    # componente do PFMI não passa no checksum) continua real e visível. A
    # causa raiz já é CREDITO_NAO_LOCALIZADO (ou o bloqueio de composição,
    # ver o laço que monta `issues` em prepare_batch); sinalizar também
    # DOC_CEDENTE_INVALIDO/DIVERGENCIA_NOME aqui seria a mesma causa
    # contada de novo, não uma informação nova para o operador.
    no_credit_resolved = analytic_credit is None and not is_composition_satellite
    nominal_suggested = (
        composition.get("nominais_sugeridos", {}).get(g.group_id) if composition else None
    )
    if nominal_suggested is not None:
        preenchimentos_automaticos.append({
            "campo": "VL_NOMINAL",
            "valor_sugerido": float(nominal_suggested),
            "regra": "PRO_RATA_VALOR_PRESENTE_COMPOSICAO",
            "exige_aprovacao_humana": True,
        })
    elif composition and composition["estado"] in ("PROPOSTA", "BLOQUEADA_DIVERGENCIA_VALOR"):
        issues.append("PREENCHER_VL_NOMINAL_COMPOSICAO")
    try:
        rate = commission_rate(p.commission_text, p.modality)
        param_issue = ""
    except ValueError as exc:
        rate, param_issue = None, str(exc)
        issues.append(param_issue)
    _, base, commission, expected, _ = _calculation(c, g) if c else (rate, None, None, None, "")
    # Divergência de VALOR com identidade já confirmada sem ambiguidade (um
    # único crédito, nome/documento já conferidos acima): preserva ambos os
    # valores para a reconciliação com tolerância em validation.py.
    # Em composição, usa o total da
    # composição inteira (mesmo valor repetido em cada componente); numa
    # linha comum, usa a aquisição esperada só desse crédito.
    divergencia_valor_analitico = None
    divergencia_valor_pfmi = None
    divergencia_valor_diferenca = None
    if composition:
        if composition["estado"] == "BLOQUEADA_DIVERGENCIA_VALOR":
            divergencia_valor_analitico = composition["total_analitico"]
            divergencia_valor_pfmi = composition["total_pfmi"]
            divergencia_valor_diferenca = composition["diferenca"]
    elif c is not None and expected is not None and g.total is not None and expected != g.total:
        divergencia_valor_analitico = expected
        divergencia_valor_pfmi = g.total
        divergencia_valor_diferenca = g.total - expected
    # The composition's principal represents the SAME real entity as the
    # matched credit (just possibly spelled differently), so - only when its
    # own ordinary match is empty - its suggested name/document default to
    # the credit's own registered values, exactly like ordinary name-based
    # matching already does. A satellite is a DIFFERENT entity: it is never
    # pre-filled from any Analítico credit (own or the composition's shared
    # one) - only from data already recorded for that same payment in the
    # PFMI itself (own component), when the document passes CPF/CNPJ
    # checksum. No cross-reference, no similarity, no proportional/averaged
    # value - if the PFMI's own beneficiary fields aren't usable, the
    # pendency (DOC_CEDENTE_INVALIDO) stays exactly as before.
    origin_doc, substitution = "ANALITICO", ""
    satellite_name_expected = None
    if is_composition_satellite:
        own_doc, own_doc_type = validate_document(p.beneficiary_document)
        own_name = str(p.beneficiary_name or "").strip()
        if own_doc_type is not None and own_name:
            name, doc = own_name, own_doc
            origin_doc = "PROPRIO_COMPONENTE_PFMI"
            # O texto do PFMI anexa o nome do titular principal entre
            # parênteses só para rastreabilidade humana ("SATELITE
            # (PRINCIPAL)"); comparar o nome próprio do satélite - já
            # comprovado pelo checksum do documento acima - contra essa
            # string inteira gera uma divergência estrutural em toda e
            # qualquer composição, nunca uma incerteza real de identidade.
            # A pendência deste satélite passa a comparar só contra a
            # própria porção dele no texto original, nunca contra o
            # titular principal entre parênteses.
            satellite_name_expected = p.cedent_name.split("(", 1)[0].strip()
            preenchimentos_automaticos.append(
                {
                    "campo": "NOME_CEDENTE/DOC_CEDENTE",
                    "valor_anterior": "",
                    "valor_sugerido": f"{own_name} / {own_doc}",
                    "regra": "PROPRIO_COMPONENTE_PFMI_DOCUMENTO_CHECKSUM_VALIDO",
                    "fonte": f"PFMI:{p.source_file}:{p.source_sheet}:linha {p.source_row}",
                    "exige_aprovacao_humana": normalize_name(own_name)
                    != normalize_name(satellite_name_expected),
                }
            )
        else:
            name, doc = "", ""
    else:
        suggestion = c or (analytic_credit if is_composition_principal else None)
        name = suggestion.cedent_name if suggestion else ""
        doc = suggestion.cedent_document if suggestion else ""
    name_expected = (
        satellite_name_expected if satellite_name_expected is not None else p.cedent_name
    )
    conex_state, conex_origins = "", ""
    if p.modality == CESSAO:
        name = name_expected = CONEXCRED_NAME
        conex = [
            pay
            for pay in file_payments
            if _is_conexcred_beneficiary(pay.beneficiary_name)
        ]
        docs = [validate_document(pay.beneficiary_document) for pay in conex]
        unique = {d for d, _ in docs}
        doc = (
            next(iter(unique)) if docs and len(unique) == 1 and all(t == 2 for _, t in docs) else ""
        )
        conex_state = "UNICO_VALIDO" if doc else "CONFLITANTE" if len(unique) > 1 else "AUSENTE"
        conex_origins = json.dumps(
            [
                dict(aba=pay.source_sheet, linha=pay.source_row, documento=pay.beneficiary_document)
                for pay in conex
            ],
            ensure_ascii=False,
        )
        if not doc:
            issues.append("CNPJ_CONEXCRED_" + conex_state)
        origin_doc = "BENEFICIARIO_CONEXCRED_PFMI" if doc else "PREENCHER_CNPJ_CONEXCRED"
        substitution = "CESSAO_DA_CESSAO_PRINCIPAL_CONEXCRED"
        # Informed once per PFMI file and replicated to every title of that
        # file. Never derived from the original signature, liquidation or
        # today's date; a missing/invalid/late value is a clear pendency,
        # never an exception, and the field stays editable/auditable.
        raw_emission = str(p.emissao_nova_cessao_texto or "").strip()
        emission_date = None
        if not raw_emission:
            issues.append("EMISSAO_NOVA_CESSAO_AUSENTE")
        else:
            try:
                emission_date = parse_date(raw_emission)
            except ValueError:
                issues.append("EMISSAO_NOVA_CESSAO_INVALIDA")
            else:
                if liquidation_date and emission_date > liquidation_date:
                    issues.append("EMISSAO_NOVA_CESSAO_POSTERIOR_LIQUIDACAO")
    elif is_lawyer_case(p.failure, lawyer_rules) or is_lawyer_case(
        failure_key(p.failure), lawyer_rules
    ):
        name = lawyer_name(name, p.failure, lawyer_rules) or lawyer_name(
            p.cedent_name, p.failure, lawyer_rules
        )
        name_expected = name or p.cedent_name
        substitution = "PAPEIS_INDEPENDENCIA_ADVOGADO_CONFERIR_INSTRUMENTO"
        issues.append("CONFERIR_ADVOGADO_DOCUMENTO_INSTRUMENTO")
    doc_type = validate_document(doc)[1]
    if doc_type is None:
        if not no_credit_resolved:
            issues.append("DOC_CEDENTE_INVALIDO")
    else:
        # TIPO_PESSOA_CEDENTE nunca decide pelo tamanho do documento: só é
        # preenchido quando o próprio documento sugerido já passou no
        # checksum completo de CPF/CNPJ acima.
        preenchimentos_automaticos.append(
            {
                "campo": "TIPO_PESSOA_CEDENTE",
                "valor_anterior": "",
                "valor_sugerido": doc_type,
                "regra": "CHECKSUM_CPF_VALIDO" if doc_type == 1 else "CHECKSUM_CNPJ_VALIDO",
                "fonte": origin_doc,
                "exige_aprovacao_humana": False,
            }
        )
    # Grafia equivalente sob documento independente confirmado: a identidade
    # já está provada de forma determinística - não pela semelhança do nome,
    # mas por DOIS documentos de fontes independentes (PFMI e Analítico),
    # cada um já validado por checksum completo, que concordam exatamente.
    # Só então a diferença de grafia (acentos/caixa já tolerados por
    # normalize_name, mais stopwords de ligação e a variante S/Z
    # documentada) deixa de ser uma divergência de identidade e passa a ser
    # só uma questão de formatação do campo final - nunca o inverso: sem
    # essa concordância documental independente, qualquer diferença de nome
    # continua exigindo APROVADO=SIM manual, sem exceção.
    document_name_confirmed = False
    name_auto_approved = False
    if (
        not is_composition_satellite
        and p.modality != CESSAO
        and not substitution
        and analytic_credit is not None
    ):
        pfmi_confirm_doc, pfmi_confirm_type = validate_document(p.beneficiary_document)
        analytic_confirm_doc, analytic_confirm_type = validate_document(
            analytic_credit.cedent_document
        )
        if (
            pfmi_confirm_type is not None
            and analytic_confirm_type is not None
            and pfmi_confirm_doc == analytic_confirm_doc
            and confirmed_document(g, analytic_credit)
        ):
            if normalize_name(name) != normalize_name(name_expected):
                name_auto_approved = True
                preenchimentos_automaticos.append(
                    {
                        "campo": "NOME_CEDENTE",
                        "valor_anterior": name,
                        "valor_sugerido": analytic_credit.cedent_name,
                        "regra": "GRAFIA_EQUIVALENTE_DOCUMENTO_INDEPENDENTE_CONFIRMADO"
                        if names_equivalent_under_confirmed_document(name, name_expected)
                        else "DOCUMENTO_EXATO_NOME_DIVERGENTE",
                        "fonte": "ANALITICO+PFMI:"
                        f"{p.source_file}:{p.source_sheet}:linha {p.source_row}",
                        "exige_aprovacao_humana": False,
                    }
                )
            name = name_expected = analytic_credit.cedent_name
            document_name_confirmed = True
            issues = [i for i in issues if i != "CREDOR_LOCALIZADO_POR_DOCUMENTO"]
    if (
        not no_credit_resolved
        and not document_name_confirmed
        and normalize_name(name) != normalize_name(name_expected)
    ):
        issues.append("DIVERGENCIA_NOME")
    source_issues = "; ".join(dict.fromkeys(i for pay in g.payments for i in pay.issues))
    issues.extend(i for pay in g.payments for i in pay.issues)
    values = dict.fromkeys(CONTROL_FIELDS, "")
    values.update(
        {
            "ID_LINHA": line_id,
            "ID_BLOCO": p.block_id,
            "ID_GRUPO": g.group_id,
            "ID_CREDITO": analytic_credit.credit_id if analytic_credit else "",
            "ID_ARQUIVO": p.file_id,
            "ORDEM_ARQUIVO": p.file_order,
            "ARQUIVO_PFMI": p.source_file,
            "HASH_PFMI": p.source_hash,
            "ABA_PFMI": p.source_sheet,
            "LINHAS_PFMI": ",".join(str(pay.source_row) for pay in g.payments),
            "ORDEM_GRUPO": g.order,
            "PAGAMENTOS_GRUPO": len(g.payments),
            "MODALIDADE": p.modality,
            "COMISSAO_TEXTO": p.commission_text,
            "PERCENTUAL_USADO": str(rate) if rate is not None else "",
            "PENDENCIAS_PARAMETROS": param_issue,
            "TOTAL_PFMI_GRUPO": g.total,
            "NOME_CEDENTE_PFMI": p.cedent_name,
            "FALENCIA_PFMI": p.failure,
            "FALENCIA_CANONICA": failure_key(p.failure),
            "DOC_BENEFICIARIO_PFMI": ", ".join(
                dict.fromkeys(pay.beneficiary_document for pay in g.payments)
            ),
            "PENDENCIAS_FONTE_PFMI": source_issues,
            "HASH_ANALITICO": analytic_credit.source_hash if analytic_credit else "",
            "ABA_ANALITICO": analytic_credit.source_sheet if analytic_credit else "",
            "LINHA_ANALITICO": analytic_credit.source_row if analytic_credit else "",
            "REFERENCIA_ANALITICO": analytic_credit.reference if analytic_credit else "",
            "NOME_CEDENTE_ANALITICO": analytic_credit.cedent_name if analytic_credit else "",
            "DOC_CEDENTE_ANALITICO": analytic_credit.cedent_document if analytic_credit else "",
            "MOTIVO_CORRESPONDENCIA": reason,
            "AQUISICAO_ORIGINAL": analytic_credit.acquisition_value if analytic_credit else None,
            "NOMINAL_ORIGINAL": analytic_credit.nominal_value if analytic_credit else None,
            "ASSINATURA_ORIGINAL": analytic_credit.signature_date if analytic_credit else None,
            "BASE_COMISSAO": base,
            "COMISSAO_CALCULADA": commission,
            "AQUISICAO_ESPERADA": expected,
            "ORIGEM_DOC_CEDENTE": origin_doc,
            "ESTADO_CNPJ_CONEXCRED": conex_state,
            "ORIGENS_CNPJ_CONEXCRED": conex_origins,
            "MOTIVO_SUBSTITUICAO": substitution,
            "NOME_ESPERADO": name_expected,
            "EVIDENCIA_BASE": json.dumps(
                [asdict(r) for r in due_index.get(failure_key(p.failure), [])], default=str
            ),
            "ORIGINAL_ANALITICO": json.dumps(c.raw_values if c else {}, default=str),
            "INCLUIR_CNAB": selected,
            "APROVADO": "",
            "STATUS": _status(issues),
            "PENDENCIAS": "; ".join(dict.fromkeys(issues)),
            "ACAO_NECESSARIA": "Corrigir parâmetro/fonte e preparar novamente"
            if param_issue or source_issues or (not c and not composition)
            else "Encontramos mais de um crédito. Selecione os créditos corretos."
            if "MULTIPLOS_CREDITOS_PRINCIPAIS" in issues
            else "Conferir referências, escolher créditos e corrigir campos finais no Excel",
            "EXIGE_NOVA_PREPARACAO": (
                "SIM" if param_issue or source_issues or (not c and not composition) else "NAO"
            ),
            "DOC_CEDENTE": digits(doc),
            "NOME_CEDENTE": name,
            "SEU_NUMERO": "",
            "NU_DOCUMENTO": "",
            "DT_VENCIMENTO": due.due_date if due else None,
            "VL_NOMINAL": nominal_suggested if composition else (c.nominal_value if c else None),
            "VL_NOMINAL_SUGERIDO": nominal_suggested,
            "DOC_SACADO": digits(due.debtor_document) if due else "",
            "NOME_SACADO": due.debtor_name if due else "",
            # Cada componente da composição usa seu próprio total do PFMI
            # (decisão registrada), nunca a aquisição do crédito inteiro.
            "VL_PRESENTE": g.total if composition else expected,
            "TIPO_PESSOA_SACADO": validate_document(due.debtor_document)[1] if due else "",
            "ENDERECO": "",
            "CEP": "",
            "TP_TITULO": 24,
            "DT_EMISSAO_TITULO": (
                emission_date
                if p.modality == CESSAO
                else analytic_credit.signature_date
                if analytic_credit
                else None
            ),
            "EMISSAO_NOVA_CESSAO_TEXTO_ORIGINAL": raw_emission if p.modality == CESSAO else "",
            "COOBRIGACAO": 2,
            "TIPO_PESSOA_CEDENTE": doc_type or "",
            "NFE": "",
            "VALOR_PAGO_TITULO": Decimal("0.00"),
            "INDEXADOR": "",
            "TAXA_INDEXADOR": "",
            "MOVIMENTO": 1,
            "DIFERENCA_PRESENTE": (
                Decimal(0)
                if composition or expected is not None
                else None
            ),
            "DIFERENCA_NOMINAL": (
                None if composition else (Decimal(0) if c and c.nominal_value is not None else None)
            ),
            "COMPOSICAO_ID": composition["composicao_id"] if composition else "",
            "COMPOSICAO_REGRA": composition["regra"] if composition else "",
            "COMPOSICAO_QTD_COMPONENTES": (
                len(composition["component_ids"]) if composition else ""
            ),
            "COMPOSICAO_PARTICIPANTES": composition["participantes"] if composition else "",
            "COMPOSICAO_TOTAL_PFMI": composition["total_pfmi"] if composition else None,
            "COMPOSICAO_TOTAL_ANALITICO": composition["total_analitico"] if composition else None,
            "COMPOSICAO_DIFERENCA_PRESENTE": composition["diferenca"] if composition else None,
            "COMPOSICAO_NOMINAL_ANALITICO": (
                composition["nominal_analitico"] if composition else None
            ),
            "COMPOSICAO_MOTIVO": composition["motivo"] if composition else "",
            "COMPOSICAO_ESTADO": composition["estado"] if composition else "",
            "COMPOSICAO_APROVADA": "",
            "COMPOSICAO_CANDIDATOS": (
                json.dumps(
                    composition["candidatos_localizacao_manual"], default=str, ensure_ascii=False
                )
                if composition and composition["candidatos_localizacao_manual"]
                else ""
            ),
            "DIVERGENCIA_VALOR_ANALITICO": divergencia_valor_analitico,
            "DIVERGENCIA_VALOR_PFMI": divergencia_valor_pfmi,
            "DIVERGENCIA_VALOR_DIFERENCA": divergencia_valor_diferenca,
            "DIVERGENCIA_VALOR_APROVADA": "",
            "DIVERGENCIA_VALOR_JUSTIFICATIVA": "",
            "PREENCHIMENTOS_AUTOMATICOS": (
                json.dumps(preenchimentos_automaticos, default=str, ensure_ascii=False)
                if preenchimentos_automaticos
                else ""
            ),
        }
    )
    for field in ("NOME_CEDENTE", "NOME_SACADO"):
        try:
            text = technical_name(values[field])
            values[field + "_TECNICO"] = text[:40]
            if len(text) > 40:
                values["ALERTAS"] += field + ": NOME_TRUNCADO_CNAB; "
        except ValueError:
            pass
    reconcile_pfmi = composition["total_pfmi"] if composition else g.total
    reconcile_analytic = composition["total_analitico"] if composition else expected
    if (
        analytic_credit is not None and reconcile_analytic is not None
        and reconcile_analytic >= 0 and reconcile_pfmi is not None and reconcile_pfmi >= 0
    ):
        reconciliation = reconcile_credit(
            composition["composicao_id"] if composition else g.group_id,
            analytic_credit.cedent_document, analytic_credit.cedent_name,
            reconcile_pfmi, reconcile_analytic,
        )
        if reconciliation.warningMessage:
            values["ALERTAS"] += reconciliation.warningMessage
            # Only a uniquely identified principal can be selected automatically.
            # Name, document, source and duplicate-credit issues remain blocking.
            if not composition and "DIVERGENCIA_VALOR" in issues and set(issues) <= {
                "DIVERGENCIA_VALOR", "SELECAO_MANUAL_NECESSARIA",
            }:
                values["INCLUIR_CNAB"] = "SIM"
                values["STATUS"] = "OK"
                values["PENDENCIAS"] = ""
                values["ACAO_NECESSARIA"] = "Conferir alerta de tolerância; TXT usará a PFMI"
    if name_auto_approved:
        audit_fill(values, "APROVADO", SYSTEM_APPROVAL,
                   "DOCUMENTO_EXATO_NOME_DIVERGENTE", "PFMI+ANALITICO")
        if values["STATUS"] == "OK":
            values["STATUS"] = "OK_COM_ALERTA_NOME"
    if composition and composition.get("auto_approved"):
        removable = {
            "DIVERGENCIA_NOME", "DIVERGENCIA_VALOR", "SELECAO_MANUAL_NECESSARIA",
            "CREDITO_NAO_LOCALIZADO", "CREDOR_LOCALIZADO_POR_DOCUMENTO",
            "COMPOSICAO_BLOQUEADA_SELECAO_MANUAL",
        }
        remaining = [i for i in issues if i not in removable]
        # Auto-approval is an auditable proposal. INCLUIR_CNAB remains the human decision.
        source = f"ANALITICO:{analytic_credit.source_sheet}:{analytic_credit.source_row}"
        for field, value in (
            ("COMPOSICAO_SELECAO_MANUAL_ABA", analytic_credit.source_sheet),
            ("COMPOSICAO_SELECAO_MANUAL_LINHA", analytic_credit.source_row),
            ("COMPOSICAO_APROVADA", SYSTEM_APPROVAL),
            ("SELECAO_MANUAL_APROVADA", SYSTEM_APPROVAL),
            ("APROVADO", SYSTEM_APPROVAL),
        ):
            audit_fill(values, field, value, "SMART_MATCH_" + composition["stp_rule"], source)
        values["INCLUIR_CNAB"] = "NAO"
        values["PENDENCIAS"] = "; ".join(dict.fromkeys(remaining))
        values["STATUS"] = _status(remaining) if remaining else "OK_COM_ALERTA_COMPOSICAO"
        values["ACAO_NECESSARIA"] = "Auditar composição e decidir INCLUIR_CNAB=SIM/NAO"
        # STP approves the deterministic nominal distribution together with the composition.
        records = json.loads(values.get("PREENCHIMENTOS_AUTOMATICOS") or "[]")
        for record in records:
            if record.get("campo") == "VL_NOMINAL":
                record["exige_aprovacao_humana"] = False
        values["PREENCHIMENTOS_AUTOMATICOS"] = json.dumps(records, ensure_ascii=False, default=str)
    elif composition and composition.get("stp_rule") in {"VALOR_AMBIGUO", "CREDITO_REUTILIZADO"}:
        values["STATUS"] = "COMPOSICAO_BLOQUEADA_SELECAO_MANUAL"
        values["PENDENCIAS"] += "; COMPOSICAO_BLOQUEADA_SELECAO_MANUAL"
        values["ALERTAS"] += "; SMART_MATCH_ABORTADO: " + composition["stp_rule"]
        values["INCLUIR_CNAB"] = "NAO"
    return PreparedRow(values, {k: values[k] for k in (*CONTROL_FIELDS, *FINAL_FIELDS)})


def _status(issues):
    for status in (
        "CONFLITO_BASE_VENCIMENTOS",
        "PREENCHER_MANUALMENTE",
        "DIVERGENCIA_VALOR",
        "DIVERGENCIA_NOME",
    ):
        if status in issues:
            return status
    return "REVISAR" if issues else "OK"
