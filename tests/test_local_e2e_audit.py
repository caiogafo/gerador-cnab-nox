from datetime import date
from types import SimpleNamespace

import pytest
import xlrd

from tests.synthetic_audit_helpers import credit_links, identity, raw_cell, reconcile, scenarios


def test_report_distinguishes_linked_credit_from_selected_credit():
    rows = [
        dict(
            ID_LINHA=f"L{i}",
            ID_CREDITO=f"C{i}",
            ORDEM_ARQUIVO=1,
            ORDEM_GRUPO=i,
            LINHA_ANALITICO=i,
            TOTAL_PFMI_GRUPO=100,
            VL_NOMINAL=150,
            VL_PRESENTE=100,
            INCLUIR_CNAB="SIM" if i == 1 else "NAO",
        )
        for i in (1, 2)
    ]
    evidence = dict(
        linhas_originais=rows,
        linhas_revisadas=rows,
        vinculos={"L1": {"indice": 0}, "L2": {"indice": 1}},
    )
    lot = {"target": {"details": [{"VL_PRESENTE": 100}] * 3}}
    public, _ = reconcile("SINTETICO", evidence, lot)
    assert public[0]["titulos_historicos_sem_vinculo"] == [3]
    assert public[0]["titulos_sem_vinculo_selecionado_na_etapa"] == [2, 3]
    assert public[0]["ordem_titulos_vinculados"] == [1]


def candidate(**changes):
    result = dict(
        ID_LINHA="L1",
        ID_CREDITO="C1",
        DOC_CEDENTE_ANALITICO="52998224725",
        FALENCIA_PFMI="FALENCIA ALFA",
        NOME_CEDENTE_ANALITICO="PESSOA SINTETICA",
        NOME_CEDENTE_PFMI="PESSOA SINTETICA",
        NOMINAL_ORIGINAL=150,
        AQUISICAO_ESPERADA=100,
        TOTAL_PFMI_GRUPO=100,
        MODALIDADE="NORMAL",
        MOTIVO_CORRESPONDENCIA="PRINCIPAL_NOME_FALENCIA_CONJUNTO_INTEGRAL",
    )
    return result | changes


def historical(**changes):
    return (
        dict(
            DOC_CEDENTE="52998224725",
            NOME_CEDENTE="PESSOA SINTETICA",
            NOME_SACADO="FALENCIA ALFA",
            VL_NOMINAL=150,
            VL_PRESENTE=100,
            DT_EMISSAO_TITULO=date(2026, 8, 1),
        )
        | changes
    )


def test_matrix_has_exactly_twenty_distinct_scenarios():
    cases = scenarios()
    assert len(cases) == len({c["id"] for c in cases}) == 20
    assert {lot: sum(c["lote"] == lot for c in cases) for lot in (1, 6, 7, 8)} == {
        1: 4,
        6: 4,
        7: 4,
        8: 8,
    }
    assert all(c["parametros_historicos_confirmados"] == (c["lote"] == 1) for c in cases)


def test_xls_serial_date_compares_as_json_date_not_serial():
    serial = xlrd.xldate.xldate_from_date_tuple((2026, 8, 14), 0)
    assert raw_cell(SimpleNamespace(ctype=xlrd.XL_CELL_DATE, value=serial), 0) == "2026-08-14"
    assert raw_cell(SimpleNamespace(ctype=xlrd.XL_CELL_EMPTY, value=""), 0) == ""


@pytest.mark.parametrize(
    "changes",
    [
        {"ID_CREDITO": ""},
        {"FALENCIA_PFMI": "OUTRA FALENCIA"},
        {"NOMINAL_ORIGINAL": 151},
        {"DOC_CEDENTE_ANALITICO": "", "NOME_CEDENTE_ANALITICO": "PESSOA (SINTETICA)"},
    ],
)
def test_history_never_creates_origin_or_silent_alias(changes):
    assert credit_links([candidate(**changes)], [historical()]) == {}


def test_duplicate_candidates_and_duplicate_historical_titles_are_ambiguous():
    row = candidate()
    assert credit_links([row, candidate(ID_LINHA="L2", ID_CREDITO="C2")], [historical()]) == {}
    assert credit_links([row], [historical(), historical()]) == {}
    assert credit_links([row], [historical()])["L1"]["indice"] == 0


def test_principal_document_is_not_an_original_credit_key():
    h = historical(
        DOC_CEDENTE="11222333000181", NOME_CEDENTE="CONEXCRED (LEGADO)", VL_PRESENTE=107.5
    )
    row = candidate(MODALIDADE="CESSAO_DA_CESSAO", AQUISICAO_ESPERADA=107.5, TOTAL_PFMI_GRUPO=107.5)
    assert credit_links([row], [h])["L1"]["evidencia"].startswith("ORIGEM_PRINCIPAL")
    # Same value with a different original person cannot establish identity.
    assert credit_links([row | {"NOME_CEDENTE_ANALITICO": "OUTRA PESSOA"}], [h]) == {}


def test_cession_never_uses_normal_historical_title_as_correction_reference():
    row = candidate(MODALIDADE="CESSAO_DA_CESSAO")
    assert credit_links([row], [historical()]) == {}


def test_identity_aliases_are_token_based():
    assert identity("ALFA C.I.A. L.T.D.A. S/A") == "ALFA CIA LTDA SA"
    assert identity("ALFALIMITADA") == "ALFALIMITADA"
