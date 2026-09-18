from dataclasses import replace
from datetime import date

import pytest

from gerador_cnab_nox.matching import _calculation, _group_payments, prepare_batch
from gerador_cnab_nox.models import PfmiData
from gerador_cnab_nox.normalize import normalize_failure
from gerador_cnab_nox.service import generate_cnab
from gerador_cnab_nox.stp import suggested_credit
from gerador_cnab_nox.workbook import write_intermediate

from .conftest import VALID_CNPJ
from .test_matching import _credit, _due, _payment
from .test_stp import _batch


def test_suggestion_fills_operational_fields_and_generates_without_edits(tmp_path):
    payment = replace(_payment("B1", "100"), failure="BANCO SANTOS")
    credit = replace(_credit("A", "100"), campaign="BANCO SANTOS S/A",
                     source_sheet="Créditos - 2026", source_row=3984)
    batch = prepare_batch(PfmiData(payments=[payment]), [credit],
                          [replace(_due(), debtor_name="BANCO SANTOS")],
                          liquidation_date=date(2026, 9, 18), first_sequence=1,
                          debtor_fallbacks={})
    row = batch.rows[0].values
    assert row["COMPOSICAO_SELECAO_MANUAL_ABA"] == "Créditos - 2026"
    assert row["COMPOSICAO_SELECAO_MANUAL_LINHA"] == 3984
    assert isinstance(row["COMPOSICAO_SELECAO_MANUAL_LINHA"], int)
    assert row["INCLUIR_CNAB"] == row["APROVADO"] == "SIM"
    assert not row["PENDENCIAS"]
    assert "AUTOMATED_MATCH_APPLIED" in row["ALERTAS"]
    path = tmp_path / "review.xlsx"
    write_intermediate(batch, path)
    assert generate_cnab(path, output_path=tmp_path / "out.txt").detail_count == 1


@pytest.mark.parametrize("case", ["wrong_doc", "wrong_failure", "wrong_amount", "duplicate",
                                  "invalid_location", "malformed"])
def test_suggestion_text_cannot_manufacture_evidence(case):
    group = _group_payments([_payment("B1", "100")])[0]
    credit = replace(_credit("A", "100"), source_sheet="Sheet1", source_row=3984)
    text = "SUGESTAO: Aba: Sheet1, Linha: 3984"
    credits = [credit]
    if case == "wrong_doc":
        credits = [replace(credit, cedent_document=VALID_CNPJ)]
    elif case == "wrong_failure":
        credits = [replace(credit, campaign="OUTRA FALENCIA")]
    elif case == "wrong_amount":
        credits = [replace(credit, acquisition_value=credit.acquisition_value + 101)]
    elif case == "duplicate":
        credits.append(replace(credit, reference="B", source_row=4000))
    elif case == "invalid_location":
        text = "Aba: Sheet1, Linha: 1"
    else:
        text = "Aba: Sheet1, Linha: nenhuma"
    assert suggested_credit(text, group, credits, _calculation, normalize_failure) is None


def test_confirmed_composition_is_included_without_manual_approval(tmp_path):
    batch = _batch()
    for row in batch.rows:
        assert row.values["INCLUIR_CNAB"] == row.values["COMPOSICAO_APROVADA"] == "SIM"
        assert row.values["APROVADO"] == "SIM"
    path = tmp_path / "review.xlsx"
    write_intermediate(batch, path)
    assert generate_cnab(path, output_path=tmp_path / "out.txt").detail_count == 2


def test_orphan_remains_unselected():
    batch = prepare_batch(PfmiData(payments=[_payment("B1", "100")]), [], [_due()],
                          liquidation_date=date(2026, 9, 18), first_sequence=1,
                          debtor_fallbacks={})
    assert batch.rows[0].values["INCLUIR_CNAB"] == "NAO"
    assert not batch.rows[0].values["APROVADO"]
