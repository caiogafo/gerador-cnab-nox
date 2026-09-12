"""Offline, read-only source diagnosis. Suggestions never alter CNAB selection."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import defaultdict
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path

from gerador_cnab_nox.matching import _calculation, _group_payments, prepare_batch
from gerador_cnab_nox.models import PfmiInput
from gerador_cnab_nox.normalize import normalize_failure, normalize_name
from gerador_cnab_nox.readers import read_analytic, read_due_base, read_pfmis


def name_variants(name):
    """Search hints only; preserve the original name and identity."""
    variants = {normalize_name(name)}
    variants.add(re.sub(r"^ESPOLIO(?: DE)?\s+", "", normalize_name(name)))
    for inner in re.findall(r"\(([^()]*)\)", name):
        variants.add(normalize_name(inner))
    return sorted(v for v in variants if v)


def suggest(group, analytic):
    """Rank exact-value candidates, explicitly including unconfirmed campaigns.

    The 0.8 text threshold is an exploratory filter, never a probability or
    an approval threshold. No subset sums or cross-block merging are performed.
    """
    if group.total is None or any(p.issues for p in group.payments):
        return []
    variants = name_variants(group.first.cedent_name)
    result = []
    for credit in analytic:
        if _calculation(credit, group)[3] != group.total:
            continue
        name = normalize_name(credit.cedent_name)
        score = max((SequenceMatcher(None, v, name).ratio() for v in variants), default=0)
        if score < 0.8:
            continue
        same_campaign = normalize_failure(credit.campaign) == normalize_failure(group.first.failure)
        result.append({
            "linha_analitico": credit.source_row,
            "similaridade_textual": round(score, 4),
            "falencia_equivalente_confirmada": same_campaign,
            "valor_exato": True,
            "acao": "CONFIRMAR_IDENTIDADE" if same_campaign else "CONFIRMAR_IDENTIDADE_E_FALENCIA",
        })
    return sorted(result, key=lambda r: (-r["similaridade_textual"], r["linha_analitico"]))


def diagnose(pfmi, analytic, due_base, liquidation_date):
    batch = prepare_batch(
        pfmi, analytic, due_base, liquidation_date=liquidation_date, first_sequence=None,
    )
    rows = defaultdict(list)
    for row in batch.rows:
        rows[row.values["ID_GRUPO"]].append(row.values)
    report_groups = []
    for group in _group_payments(pfmi.payments):
        current = rows[group.group_id]
        selected = [r for r in current if r["INCLUIR_CNAB"] == "SIM"]
        report_groups.append({
            "grupo": group.order,
            "aba_pfmi": group.first.source_sheet,
            "linhas_pfmi": [p.source_row for p in group.payments],
            "pagamentos": len(group.payments),
            "total": str(group.total) if group.total is not None else None,
            "selecionado_atualmente": bool(selected),
            "linhas_analitico_selecionadas": [r["LINHA_ANALITICO"] for r in selected],
            "pendencias": sorted({
                issue.strip() for r in current for issue in r["PENDENCIAS"].split(";")
                if issue.strip()
            }),
            "sugestoes_nao_aprovadas": [] if selected else suggest(group, analytic),
        })
    return {
        "escopo": "Diagnostico de cruzamento; nao valida nem gera CNAB",
        "liquidacao_hipotese": str(liquidation_date),
        "primeira_sequencia": None,
        "pagamentos": len(pfmi.payments),
        "linhas_nao_reconhecidas": len(pfmi.unknown_rows),
        "grupos": len(report_groups),
        "grupos_selecionados": sum(g["selecionado_atualmente"] for g in report_groups),
        "grupos_com_sugestao": sum(bool(g["sugestoes_nao_aprovadas"]) for g in report_groups),
        "detalhes": report_groups,
    }


def fingerprint(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def markdown(report):
    lines = [
        "# Diagnóstico de reconciliação",
        "",
        "Análise local. Nenhum TXT gerado; sugestões não são títulos aprovados.",
        "Modalidade e liquidação são hipóteses explícitas de execução.",
        "Primeira sequência não informada. Não foi feita validação completa de geração.",
        "",
        f"Pagamentos: {report['pagamentos']}; grupos: {report['grupos']}; "
        f"grupos pré-selecionados: {report['grupos_selecionados']}; "
        f"grupos adicionais com sugestões: {report['grupos_com_sugestao']}.",
        "",
        "| Grupo | Linhas PFMI | Pré-selecionado | Linhas Analítico / sugestões | Pendências |",
        "|---|---|---|---|---|",
    ]
    for g in report["detalhes"]:
        refs = g["linhas_analitico_selecionadas"] or [
            s["linha_analitico"] for s in g["sugestoes_nao_aprovadas"]
        ]
        lines.append(
            f"| {g['grupo']} | {','.join(map(str, g['linhas_pfmi']))} | "
            f"{'Sim' if g['selecionado_atualmente'] else 'Não'} | "
            f"{', '.join(map(str, refs)) or 'Sem sugestão'} | "
            f"{'; '.join(g['pendencias']) or 'Sem pendência nesta preparação'} |"
        )
    lines.extend([
        "", "Sugestões exigem revisão da identidade e, quando indicado no JSON, da falência.",
        "Similaridade textual não mede probabilidade de identidade. Valores são comparados",
        "em centavos pela regra de modalidade. Blocos separados não foram somados.",
        "Fontes preservadas por SHA-256 antes e depois da análise.",
    ])
    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pfmi", required=True, type=Path)
    parser.add_argument("--analytic", required=True, type=Path)
    parser.add_argument("--base", required=True, type=Path)
    parser.add_argument("--date", required=True, type=date.fromisoformat)
    parser.add_argument("--modality", required=True, choices=["NORMAL", "CESSAO_DA_CESSAO"])
    parser.add_argument("--commission", default="15")
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    if args.output.exists():
        parser.error("O destino deve ser novo; nenhuma saída é sobrescrita.")
    sources = {"pfmi": args.pfmi, "analitico": args.analytic, "base": args.base}
    hashes = {role: fingerprint(path) for role, path in sources.items()}
    report = diagnose(
        read_pfmis([PfmiInput(args.pfmi, args.modality, args.commission)]),
        read_analytic(args.analytic), read_due_base(args.base), args.date,
    )
    if hashes != {role: fingerprint(path) for role, path in sources.items()}:
        raise RuntimeError("As fontes mudaram durante a análise; resultado descartado.")
    report["sha256_fontes"] = hashes
    report["modalidade_hipotese"] = args.modality
    report["comissao_parametro"] = args.commission
    args.output.mkdir(parents=True, exist_ok=False)
    (args.output / "diagnostico.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    (args.output / "RELATORIO.md").write_text(markdown(report), encoding="utf-8")
    print(json.dumps({k: report[k] for k in (
        "pagamentos", "grupos", "grupos_selecionados", "grupos_com_sugestao",
    )}))


if __name__ == "__main__":
    main()
