"""Generic, data-free helpers extracted for the synthetic audit tests.

Pure functions only: normalization, hashing/JSON-saving utilities and a fixed
combinatorial scenario matrix. Nothing here reads or references any real
batch, historical file or operational corpus.
"""

from __future__ import annotations

import hashlib
import itertools
import json
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from gerador_cnab_nox.models import CESSAO, NORMAL

CENT = Decimal("0.01")
SELECTED_LOTS = (1, 6, 7, 8)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def encoded(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, Path):
        return str(value)
    raise TypeError(type(value).__name__)


def save(path, value):
    with Path(path).open("x", encoding="utf-8") as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2, default=encoded)
    Path(path).chmod(0o600)


def cents(value):
    if value is None or value == "":
        return None
    return int(Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP) * 100)


def text_name(value):
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    return " ".join("".join(c for c in text if not unicodedata.combining(c)).split())


def identity(value):
    text = text_name(value)
    for pattern, replacement in (
        (r"\bL\s*\.\s*T\s*\.\s*D\s*\.\s*A\s*\.?", "LTDA"),
        (r"\bC\s*\.\s*I\s*\.\s*A\s*\.?", "CIA"),
        (r"\bS\s*[./]\s*A\s*\.?", "SA"),
    ):
        text = re.sub(pattern, replacement + " ", text)
    words = re.sub(r"[^A-Z0-9]+", " ", text).split()
    result, n = [], 0
    while n < len(words):
        if words[n : n + 2] == ["SOCIEDADE", "ANONIMA"]:
            result.append("SA")
            n += 2
        else:
            result.append({"LIMITADA": "LTDA", "COMPANHIA": "CIA"}.get(words[n], words[n]))
            n += 1
    return " ".join(result)


def failure(value):
    return identity(value)


def doc(value):
    return re.sub(r"\D", "", str(value or ""))


def raw_cell(cell, datemode):
    import xlrd

    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode).date().isoformat()
    return cell.value


def scenarios():
    result = []
    for lot in SELECTED_LOTS:
        modes = (
            [(NORMAL, CESSAO)]
            if lot == 1
            else list(itertools.product((NORMAL, CESSAO), repeat=2 if lot == 8 else 1))
        )
        for mode, a, b in itertools.product(
            modes, ("raiz", "lote01"), ("raiz", "lote01") if lot == 1 else ("raiz",)
        ):
            tag = "".join("N" if m == NORMAL else "C" for m in mode)
            result.append(
                dict(
                    id=f"T{lot:02}_{tag}_A{a}_B{b}",
                    lote=lot,
                    modalidades=list(mode),
                    analitico=a,
                    base=b,
                    parametros_historicos_confirmados=lot == 1,
                )
            )
    if len(result) != 20:
        raise ValueError("Matriz sintética esperada exige 20 combinações")
    return result


def credit_links(rows, historical):
    """Conservative audit crosswalk; the production candidate origin is mandatory.

    A unique document+campaign+nominal is identity evidence. Cession additionally
    needs the original full PFMI name, original credit and individual commission;
    the principal's historical document never identifies the original credit.
    """
    proposed = {}
    for v in rows:
        if not v["ID_CREDITO"]:
            continue
        same_failure = [
            i
            for i, h in enumerate(historical)
            if failure(v["FALENCIA_PFMI"]) == failure(h["NOME_SACADO"])
        ]
        hits = []
        reason = "DOCUMENTO_ORIGINAL_FALENCIA_NOMINAL_UNICOS"
        if v["MODALIDADE"] == NORMAL:
            hits = [
                i
                for i in same_failure
                if doc(v["DOC_CEDENTE_ANALITICO"])
                and doc(v["DOC_CEDENTE_ANALITICO"]) == doc(historical[i]["DOC_CEDENTE"])
                and cents(v["NOMINAL_ORIGINAL"]) == cents(historical[i]["VL_NOMINAL"])
            ]
        elif (
            v["MODALIDADE"] == CESSAO
            and identity(v["NOME_CEDENTE_ANALITICO"]) == identity(v["NOME_CEDENTE_PFMI"])
            and v["MOTIVO_CORRESPONDENCIA"] == "PRINCIPAL_NOME_FALENCIA_CONJUNTO_INTEGRAL"
        ):
            hits = [
                i
                for i in same_failure
                if text_name(historical[i]["NOME_CEDENTE"]).startswith("CONEXCRED")
                and cents(v["NOMINAL_ORIGINAL"]) == cents(historical[i]["VL_NOMINAL"])
                and cents(v["AQUISICAO_ESPERADA"]) == cents(historical[i]["VL_PRESENTE"])
                and cents(v["TOTAL_PFMI_GRUPO"]) == cents(historical[i]["VL_PRESENTE"])
            ]
            reason = "ORIGEM_PRINCIPAL_NOME_FALENCIA_NOMINAL_COMISSAO"
        if len(hits) == 1:
            proposed[v["ID_LINHA"]] = dict(
                indice=hits[0], evidencia=reason, id_credito=v["ID_CREDITO"]
            )
    used = Counter(v["indice"] for v in proposed.values())
    credit_used = Counter(v["id_credito"] for v in proposed.values())
    return {
        k: v
        for k, v in proposed.items()
        if used[v["indice"]] == 1 and credit_used[v["id_credito"]] == 1
    }


def reconcile(case, private, lot):
    """Account for partial selections without claiming unmatched records are extras."""
    historical = lot["target"]["details"]
    links = private["vinculos"]
    public, detailed = [], []
    for stage, key in (("original", "linhas_originais"), ("revisada", "linhas_revisadas")):
        rows = private[key]
        groups = defaultdict(list)
        for row in rows:
            groups[(row["ORDEM_ARQUIVO"], row["ORDEM_GRUPO"])].append(row)
        group_report, private_groups = [], []
        pfmi_total = 0
        for (file_no, group_no), candidates in groups.items():
            targets = {cents(v["TOTAL_PFMI_GRUPO"]) for v in candidates}
            if len(targets) != 1:
                raise RuntimeError("Total de grupo inconsistente nas evidências")
            target = targets.pop()
            pfmi_total += target
            selected = [v for v in candidates if v["INCLUIR_CNAB"] == "SIM"]
            present = [
                cents(v["VL_PRESENTE"]) for v in selected if v.get("VL_PRESENTE") is not None
            ]
            nominal = [cents(v["VL_NOMINAL"]) for v in selected if v.get("VL_NOMINAL") is not None]
            origins = sum(bool(v.get("ID_CREDITO")) for v in candidates)
            info = dict(
                arquivo=file_no,
                grupo=group_no,
                candidatos_com_origem=origins,
                selecionados=len(selected),
                aquisicoes_ausentes=len(selected) - len(present),
                delta_aquisicao_selecionada_pfmi_centavos=sum(present) - target,
                titulos_historicos_vinculados=[
                    links[v["ID_LINHA"]]["indice"] + 1 for v in selected if v["ID_LINHA"] in links
                ],
                acao=(
                    "RECONCILIAR_ORIGEM_E_PREPARAR_NOVAMENTE"
                    if not origins
                    else (
                        "REVISAR_IDENTIDADE_E_SELECAO"
                        if not selected
                        else "CONFERIR_PENDENCIAS_DE_CAMPOS_E_RECONCILIACAO"
                    )
                ),
                limite="Seleção parcial de Excel bloqueado; não constitui lote válido",
            )
            group_report.append(info)
            private_groups.append(
                dict(
                    **info,
                    pfmi_centavos=target,
                    aquisicao_centavos=sum(present),
                    nominal_centavos=sum(nominal),
                )
            )
        selected = sorted(
            (v for v in rows if v["INCLUIR_CNAB"] == "SIM"),
            key=lambda v: (v["ORDEM_ARQUIVO"], v["ORDEM_GRUPO"], v["LINHA_ANALITICO"]),
        )
        linked_order = [
            links[v["ID_LINHA"]]["indice"] + 1 for v in selected if v["ID_LINHA"] in links
        ]
        ids = Counter(v["ID_CREDITO"] for v in selected if v.get("ID_CREDITO"))
        public.append(
            dict(
                cenario=case,
                etapa=stage,
                grupos=group_report,
                delta_total_pfmi_historico_centavos=pfmi_total
                - sum(cents(h["VL_PRESENTE"]) for h in historical),
                ordem_titulos_vinculados=linked_order,
                ordem_relativa_coincide=linked_order == sorted(linked_order),
                creditos_reutilizados=sum(n - 1 for n in ids.values() if n > 1),
                titulos_historicos_sem_vinculo=sorted(
                    set(range(1, len(historical) + 1))
                    - {link["indice"] + 1 for link in links.values()}
                ),
                titulos_sem_vinculo_selecionado_na_etapa=sorted(
                    set(range(1, len(historical) + 1)) - set(linked_order)
                ),
                selecionados_sem_vinculo=len(selected) - len(linked_order),
                limite="Sem vínculo é inconclusivo; não significa título ausente ou extra no TXT",
            )
        )
        detailed.append(dict(cenario=case, etapa=stage, grupos=private_groups))
    return public, detailed
