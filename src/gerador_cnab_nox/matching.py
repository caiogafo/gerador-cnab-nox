from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict
from dataclasses import asdict, dataclass
from decimal import Decimal

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
    normalize_failure,
    normalize_name,
    parse_date,
    technical_name,
    validate_document,
)

CONEXCRED_NAME = "CONEXCRED INTERMEDIACAO"


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


def prepare_batch(
    pfmi: PfmiData,
    analytic: list[AnalyticCredit],
    due_base: list[DueRecord],
    *,
    liquidation_date,
    first_sequence,
    failure_aliases=None,
    lawyer_rules=None,
) -> PreparedBatch:
    lawyer_rules = lawyer_rules or {}
    failure_key = resolver(failure_aliases or {})
    groups = _group_payments(pfmi.payments, failure_key)
    due_index = _index_due_base(due_base, failure_key)
    calculations, primary, candidates, eligible = {}, {}, {}, {}
    for g in groups:
        main = [
            c
            for c in analytic
            if normalize_name(c.cedent_name) == normalize_name(g.first.cedent_name)
            and failure_key(c.campaign) == failure_key(g.first.failure)
        ]
        primary[g.group_id] = main
        candidates[g.group_id] = list(main)
        same_failure = [
            c
            for c in analytic
            if failure_key(c.campaign) == failure_key(g.first.failure)
        ]
        for c in same_failure:
            calculations[g.group_id, c.credit_id] = _calculation(c, g)
        docs = [validate_document(c.cedent_document) for c in main]
        expected = [calculations[g.group_id, c.credit_id][3] for c in main]
        eligible[g.group_id] = (
            len(main) == 1
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
        credits = sorted(candidates[g.group_id], key=lambda c: c.source_row)
        for c in credits or [None]:
            issues = list(due_issues)
            if not auto:
                issues.append("SELECAO_MANUAL_NECESSARIA")
            if status:
                issues.append(status)
            reason = "PRINCIPAL_NOME_FALENCIA_CONJUNTO_INTEGRAL"
            if c is None:
                issues.append("CREDITO_NAO_LOCALIZADO")
                reason = "CREDITO_NAO_LOCALIZADO"
            else:
                calc = calculations[g.group_id, c.credit_id]
                if calc[4]:
                    issues.append(calc[4])
                if c.credit_id not in main_ids:
                    reason = "ALTERNATIVA_VALOR_EXATO_ESCOLHA_MANUAL"
                    issues.append("DIVERGENCIA_NOME")
                elif not eligible[g.group_id]:
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
            prepared.append(
                _make_row(
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
                )
            )
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


def _make_row(g, c, due, due_index, issues, reason, selected, file_payments,
              failure_key=normalize_failure, lawyer_rules=None, liquidation_date=None):
    lawyer_rules = lawyer_rules or {}
    p = g.first
    credit_id = c.credit_id if c else ""
    line_id = hashlib.sha256(f"{g.group_id}|{credit_id}".encode()).hexdigest()[:20].upper()
    try:
        rate = commission_rate(p.commission_text, p.modality)
        param_issue = ""
    except ValueError as exc:
        rate, param_issue = None, str(exc)
        issues.append(param_issue)
    _, base, commission, expected, _ = _calculation(c, g) if c else (rate, None, None, None, "")
    name = c.cedent_name if c else ""
    doc = c.cedent_document if c else ""
    name_expected = p.cedent_name
    origin_doc, substitution = "ANALITICO", ""
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
    if validate_document(doc)[1] is None:
        issues.append("DOC_CEDENTE_INVALIDO")
    if normalize_name(name) != normalize_name(name_expected):
        issues.append("DIVERGENCIA_NOME")
    source_issues = "; ".join(dict.fromkeys(i for pay in g.payments for i in pay.issues))
    issues.extend(i for pay in g.payments for i in pay.issues)
    values = dict.fromkeys(CONTROL_FIELDS, "")
    values.update(
        {
            "ID_LINHA": line_id,
            "ID_BLOCO": p.block_id,
            "ID_GRUPO": g.group_id,
            "ID_CREDITO": credit_id,
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
            "HASH_ANALITICO": c.source_hash if c else "",
            "ABA_ANALITICO": c.source_sheet if c else "",
            "LINHA_ANALITICO": c.source_row if c else "",
            "REFERENCIA_ANALITICO": c.reference if c else "",
            "NOME_CEDENTE_ANALITICO": c.cedent_name if c else "",
            "DOC_CEDENTE_ANALITICO": c.cedent_document if c else "",
            "MOTIVO_CORRESPONDENCIA": reason,
            "AQUISICAO_ORIGINAL": c.acquisition_value if c else None,
            "NOMINAL_ORIGINAL": c.nominal_value if c else None,
            "ASSINATURA_ORIGINAL": c.signature_date if c else None,
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
            if param_issue or source_issues or not c
            else "Encontramos mais de um crédito. Selecione os créditos corretos."
            if "MULTIPLOS_CREDITOS_PRINCIPAIS" in issues
            else "Conferir referências, escolher créditos e corrigir campos finais no Excel",
            "EXIGE_NOVA_PREPARACAO": "SIM" if param_issue or source_issues or not c else "NAO",
            "DOC_CEDENTE": digits(doc),
            "NOME_CEDENTE": name,
            "SEU_NUMERO": "",
            "NU_DOCUMENTO": "",
            "DT_VENCIMENTO": due.due_date if due else None,
            "VL_NOMINAL": c.nominal_value if c else None,
            "DOC_SACADO": digits(due.debtor_document) if due else "",
            "NOME_SACADO": due.debtor_name if due else "",
            "VL_PRESENTE": expected,
            "TIPO_PESSOA_SACADO": validate_document(due.debtor_document)[1] if due else "",
            "ENDERECO": "",
            "CEP": "",
            "TP_TITULO": 24,
            "DT_EMISSAO_TITULO": (
                emission_date if p.modality == CESSAO else c.signature_date if c else None
            ),
            "EMISSAO_NOVA_CESSAO_TEXTO_ORIGINAL": raw_emission if p.modality == CESSAO else "",
            "COOBRIGACAO": 2,
            "TIPO_PESSOA_CEDENTE": validate_document(doc)[1] or "",
            "NFE": "",
            "VALOR_PAGO_TITULO": Decimal("0.00"),
            "INDEXADOR": "",
            "TAXA_INDEXADOR": "",
            "MOVIMENTO": 1,
            "DIFERENCA_PRESENTE": Decimal(0) if expected is not None else None,
            "DIFERENCA_NOMINAL": Decimal(0) if c and c.nominal_value is not None else None,
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
