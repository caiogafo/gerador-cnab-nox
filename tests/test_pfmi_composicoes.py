"""Composições: vários pagamentos do PFMI para um único crédito do Analítico.

Todos os dados são fictícios. A composição é sempre uma proposta exploratória
(matching.py::_composition_candidates); nunca confirma identidade nem preenche
INCLUIR_CNAB sozinha. VL_NOMINAL individual recebe sugestão matemática proporcional
quando os totais permitem, sujeita à revisão do operador no Excel. A geração de CNAB só
ocorre com o grupo inteiro aprovado (COMPOSICAO_APROVADA=SIM), incluído
(INCLUIR_CNAB=SIM) e com VL_NOMINAL válido em cada linha - ver
validate_for_generation em validation.py.
"""

from __future__ import annotations

import json
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from gerador_cnab_nox.errors import ValidationError
from gerador_cnab_nox.gui import _composition_message
from gerador_cnab_nox.matching import prepare_batch
from gerador_cnab_nox.models import AnalyticCredit, DueRecord, PfmiData, PfmiPayment
from gerador_cnab_nox.service import generate_cnab
from gerador_cnab_nox.validation import validate_for_generation
from gerador_cnab_nox.workbook import read_intermediate, write_intermediate

from .conftest import VALID_CNPJ, VALID_CPF

FALENCIA = "FALENCIA FICTICIA UM"
PRINCIPAL = "PRINCIPAL FICTICIO UM"
REPRESENTANTE = f"REPRESENTANTE FICTICIO ({PRINCIPAL})"
ESCRITORIO = f"ESCRITORIO FICTICIO ({PRINCIPAL})"


@pytest.mark.parametrize(
    ("nominal", "principal_expected", "satellite_expected"),
    [("1.00", "0.34", "0.33"), ("2.00", "0.66", "0.67")],
)
def test_composition_suggested_nominal_exact_sum(
    nominal, principal_expected, satellite_expected,
):
    # Principal deliberately follows a satellite in PFMI order.
    batch = _prepare(
        [_pay(REPRESENTANTE, "1.00"), _pay(PRINCIPAL, "1.00"), _pay(ESCRITORIO, "1.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="3.00", nominal=nominal)],
    )
    values = [r.values for r in batch.rows]
    assert sum((v["VL_NOMINAL_SUGERIDO"] for v in values), Decimal(0)) == Decimal(nominal)
    for v in values:
        expected = principal_expected if v["NOME_CEDENTE_PFMI"] == PRINCIPAL else satellite_expected
        assert v["VL_NOMINAL_SUGERIDO"] == Decimal(expected)
        assert v["VL_NOMINAL"] == v["VL_NOMINAL_SUGERIDO"]
        assert "PREENCHER_VL_NOMINAL_COMPOSICAO" not in v["PENDENCIAS"]
        entry = next(e for e in json.loads(v["PREENCHIMENTOS_AUTOMATICOS"])
                     if e["campo"] == "VL_NOMINAL")
        assert entry == {
            "campo": "VL_NOMINAL",
            "valor_sugerido": float(Decimal(expected)),
            "regra": "PRO_RATA_VALOR_PRESENTE_COMPOSICAO",
            "exige_aprovacao_humana": True,
        }


@pytest.mark.parametrize("approval", [None, "NAO", "TALVEZ"])
def test_composition_suggested_nominal_requires_explicit_human_approval(tmp_path, approval):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL)],
    )
    edits = _fully_approved_edits(batch.rows, [300, 200])
    for row in batch.rows:
        # Keep the prefilled nominal; all other fields are valid.
        edits[row.values["ID_LINHA"]].pop("VL_NOMINAL")
    edits[batch.rows[-1].values["ID_LINHA"]]["COMPOSICAO_APROVADA"] = approval
    path = _write_and_edit(tmp_path, batch, edits)
    target = tmp_path / "blocked.txt"
    with pytest.raises(ValidationError, match="COMPOSICAO_APROVADA"):
        generate_cnab(path, output_path=target)
    assert not target.exists()


@pytest.mark.parametrize("delta", ["0.00", "0.01", "-0.01"])
def test_composition_manual_override_validates_exact_sum(tmp_path, delta):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL)],
    )
    edits = _fully_approved_edits(batch.rows, [Decimal("100"), Decimal("400") + Decimal(delta)])
    path = _write_and_edit(tmp_path, batch, edits)
    target = tmp_path / "out.txt"
    if Decimal(delta):
        with pytest.raises(ValidationError, match="soma dos nominais"):
            generate_cnab(path, output_path=target)
        assert not target.exists()
    else:
        result = generate_cnab(path, output_path=target)
        assert result.detail_count == 2
        assert result.total_nominal == Decimal("500.00")


def test_composition_suggested_nominal_generates_after_approval_without_override(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL)],
    )
    edits = _fully_approved_edits(batch.rows, [300, 200])
    for fields in edits.values():
        fields.pop("VL_NOMINAL")
    path = _write_and_edit(tmp_path, batch, edits)
    result = generate_cnab(path, output_path=tmp_path / "approved.txt")
    assert result.detail_count == 2
    assert result.total_nominal == Decimal("500.00")


def test_composition_suggested_nominal_protection_and_roundtrip(tmp_path):
    from openpyxl import load_workbook

    from gerador_cnab_nox.errors import InputFileError
    from gerador_cnab_nox.models import IMMUTABLE_FIELDS

    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL)],
    )
    path = _write_and_edit(tmp_path, batch, {})
    book = load_workbook(path)
    sheet = book["CREDITOS"]
    header = {c.value: c.column for c in sheet[1]}
    assert "VL_NOMINAL_SUGERIDO" in IMMUTABLE_FIELDS
    assert sheet.cell(2, header["VL_NOMINAL_SUGERIDO"]).protection.locked
    assert not sheet.cell(2, header["VL_NOMINAL"]).protection.locked
    assert sheet.cell(2, header["VL_NOMINAL"]).value == 300
    assert read_intermediate(path).rows[0].originals["VL_NOMINAL_SUGERIDO"] == 300
    sheet.cell(2, header["VL_NOMINAL_SUGERIDO"]).value = 301
    book.save(path)
    book.close()
    with pytest.raises((InputFileError, ValidationError)):
        validate_for_generation(read_intermediate(path))


@pytest.mark.parametrize("field,value", [
    ("total_analitico", Decimal("0")),
    ("nominal_analitico", Decimal("0")),
    ("nominal_analitico", None),
])
def test_composition_without_positive_totals_has_no_suggestion(field, value):
    from gerador_cnab_nox.matching import _group_payments, _suggest_composition_nominals

    groups = _group_payments([_pay(PRINCIPAL, "60"), _pay(REPRESENTANTE, "40")])
    composition = {
        "total_analitico": Decimal("100"), "nominal_analitico": Decimal("500"),
        "component_ids": [g.group_id for g in groups], field: value,
    }
    assert _suggest_composition_nominals(composition, {g.group_id: g for g in groups}) == {}


def _pay(cedent, value, *, failure=FALENCIA, block="B0001", doc=VALID_CPF):
    return PfmiPayment(
        block_id=block,
        source_row=1,
        title="FUNDO",
        failure=failure,
        cedent_name=cedent,
        beneficiary_name="BENEFICIARIO FICTICIO",
        beneficiary_document=doc,
        operation_value=Decimal(value),
        beneficiary_value=Decimal(value),
    )


def _credit(reference, cedent, *, failure=FALENCIA, acquisition="100.00", nominal="500.00",
            doc=VALID_CPF, signature=date(2026, 9, 1), source_row=2, source_sheet=""):
    return AnalyticCredit(
        source_row=source_row,
        reference=reference,
        cedent_document=doc,
        cedent_name=cedent,
        campaign=failure,
        nominal_value=Decimal(nominal),
        acquisition_value=Decimal(acquisition),
        signature_date=signature,
        source_sheet=source_sheet,
    )


def _due(failure=FALENCIA):
    return DueRecord(2, failure, VALID_CNPJ, date(2030, 1, 1))


def _prepare(payments, credits, *, due_records=None, first_sequence=100):
    return prepare_batch(
        PfmiData(payments=list(payments)),
        list(credits),
        list(due_records) if due_records is not None else [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=first_sequence,
    )


def _composition_rows(batch, composicao_id):
    return [r for r in batch.rows if r.values["COMPOSICAO_ID"] == composicao_id]


# 1. Composição exata com dois componentes -----------------------------------
def test_exact_composition_with_two_components_is_proposed():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    assert len(batch.compositions) == 1
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert comp["total_pfmi"] == Decimal("100.00")
    assert comp["diferenca"] == Decimal("0.00")
    assert isinstance(comp["total_pfmi"], Decimal) and isinstance(comp["diferenca"], Decimal)
    assert PRINCIPAL in comp["participantes"] and REPRESENTANTE in comp["participantes"]
    rows = _composition_rows(batch, comp["composicao_id"])
    assert len(rows) == 2
    assert all(r.values["COMPOSICAO_ESTADO"] == "PROPOSTA" for r in rows)
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in rows)  # nunca automático
    assert all(r.values["COMPOSICAO_APROVADA"] == "" for r in rows)


# 2. Composição exata com três componentes ------------------------------------
def test_exact_composition_with_three_components_is_proposed():
    batch = _prepare(
        [_pay(PRINCIPAL, "40.00"), _pay(REPRESENTANTE, "20.00"), _pay(ESCRITORIO, "10.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="70.00")],
    )
    assert len(batch.compositions) == 1
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert comp["total_pfmi"] == Decimal("70.00")
    rows = _composition_rows(batch, comp["composicao_id"])
    assert len(rows) == 3
    assert comp["participantes"].count(";") == 2


# 3. Composição com mais componentes relacionados (quatro) --------------------
def test_composition_accepts_more_than_three_components():
    fourth = f"QUARTO FICTICIO ({PRINCIPAL})"
    batch = _prepare(
        [
            _pay(PRINCIPAL, "10.00"),
            _pay(REPRESENTANTE, "10.00"),
            _pay(ESCRITORIO, "10.00"),
            _pay(fourth, "10.00"),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="40.00")],
    )
    assert len(batch.compositions) == 1
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert len(comp["component_ids"]) == 4


# 4. Diferença de R$ 0,01 bloqueia --------------------------------------------
def test_one_cent_difference_blocks_composition():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.01")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_DIVERGENCIA_VALOR"
    assert comp["diferenca"] == Decimal("-0.01")
    rows = _composition_rows(batch, comp["composicao_id"])
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in rows)


# 5. Valores iguais em falências diferentes nunca compõem ---------------------
def test_equal_values_in_different_failures_never_compose():
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00", failure="FALENCIA FICTICIA UM"),
            _pay(REPRESENTANTE, "40.00", failure="FALENCIA FICTICIA DOIS"),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    assert batch.compositions == []


# 6. Dois créditos possíveis no Analítico bloqueia ----------------------------
def test_two_possible_credits_block_composition():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [
            _credit("REF-1", PRINCIPAL, acquisition="100.00"),
            _credit("REF-2", PRINCIPAL, acquisition="100.00"),
        ],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_MULTIPLOS_CREDITOS"


# 7. Referência entre parênteses sem relação válida ---------------------------
def test_parenthetical_reference_without_valid_relation_composes_nothing():
    lonely = "SATELITE FICTICIO (NINGUEM RELACIONADO)"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(lonely, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    assert batch.compositions == []


# 8. Grupo principal ausente ---------------------------------------------------
def test_missing_principal_group_composes_nothing():
    batch = _prepare(
        [_pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="40.00")],
    )
    assert batch.compositions == []


# 9. Composições sobrepostas (referência ambígua) -----------------------------
def test_ambiguous_reference_blocks_overlapping_composition():
    other_principal = "OUTRO PRINCIPAL FICTICIO"
    # "REPRESENTANTE FICTICIO" satelite references a name that matches TWO
    # different principals in the same file/falência: ambíguo, nunca composto.
    ambiguous_satellite = "REPRESENTANTE FICTICIO (PRINCIPAL FICTICIO)"
    principal_a = "PRINCIPAL FICTICIO"
    principal_b = "PRINCIPAL FICTICIO"  # mesmo nome normalizado -> ambiguidade
    batch = _prepare(
        [
            _pay(principal_a, "10.00", block="B0001"),
            _pay(principal_b, "10.00", block="B0002"),
            _pay(ambiguous_satellite, "20.00", block="B0003"),
        ],
        [_credit("REF-1", principal_a, acquisition="20.00")],
    )
    assert other_principal not in [c["participantes"] for c in batch.compositions]
    assert all(c["estado"] == "BLOQUEADA_REFERENCIA_AMBIGUA" for c in batch.compositions) or (
        batch.compositions == []
    )


# 10. Reutilização do crédito em matching individual e composição ------------
def test_credit_reused_by_unrelated_group_is_flagged_on_both_sides():
    # Um terceiro grupo sem qualquer relação estrutural com a composição tem,
    # por coincidência, o mesmo total do crédito já usado pela composição: a
    # proteção existente de "crédito candidato em múltiplos grupos" (uses),
    # não alterada por esta funcionalidade, continua sinalizando o conflito.
    terceiro = "TERCEIRO SEM RELACAO FICTICIO"
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00", block="B0001"),
            _pay(REPRESENTANTE, "40.00", block="B0001"),
            _pay(terceiro, "100.00", block="B0002"),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"  # a composição em si continua exata
    terceiro_row = next(r for r in batch.rows if r.values["NOME_CEDENTE_PFMI"] == terceiro)
    assert terceiro_row.values["ID_CREDITO"] == comp["credito_referencia"] or (
        "CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS" in terceiro_row.values["PENDENCIAS"]
    )
    principal_row = next(
        r
        for r in batch.rows
        if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL and r.values["ID_CREDITO"]
    )
    assert "CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS" in principal_row.values["PENDENCIAS"]
    assert terceiro_row.values["INCLUIR_CNAB"] == "NAO"
    assert principal_row.values["INCLUIR_CNAB"] == "NAO"


# 11. Similaridade de nome nunca confirma composição --------------------------
def test_near_miss_name_never_composes():
    near_miss = "REPRESENTANTE FICTICIO (PRINCIPAL FICTICIO)"  # falta "UM"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(near_miss, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    assert batch.compositions == []


# 11b. Principal reconhecido por CPF/CNPJ exato quando o nome diverge --------
def test_principal_matches_by_exact_document_when_name_diverges():
    outro_nome = "PESSOA CADASTRADA DIFERENTE NO ANALITICO"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc=VALID_CPF), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", outro_nome, acquisition="100.00", doc=VALID_CPF)],
    )
    assert len(batch.compositions) == 1
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert comp["total_pfmi"] == Decimal("100.00")
    assert "documento" in comp["motivo"].lower()


def test_document_match_never_uses_similar_document_as_confirmation():
    # CNPJ diferente (também válido) não deve casar por "parecido": só
    # igualdade exata de dígitos é aceita.
    outro_nome = "PESSOA CADASTRADA DIFERENTE NO ANALITICO"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc=VALID_CPF), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", outro_nome, acquisition="100.00", doc=VALID_CNPJ)],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_CREDITO_NAO_LOCALIZADO"


def test_document_match_blocks_when_ambiguous():
    outro_nome_a = "PESSOA CADASTRADA A"
    outro_nome_b = "PESSOA CADASTRADA B"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc=VALID_CPF), _pay(REPRESENTANTE, "40.00")],
        [
            _credit("REF-A", outro_nome_a, acquisition="100.00", doc=VALID_CPF),
            _credit("REF-B", outro_nome_b, acquisition="100.00", doc=VALID_CPF),
        ],
    )
    comp = batch.compositions[0]
    # Um documento repetido em dois créditos nunca vira candidato automático
    # (evita popular uma lista ambígua "por acidente"); some como se nenhum
    # crédito tivesse sido localizado, igual a uma ausência real de documento.
    assert comp["estado"] == "BLOQUEADA_CREDITO_NAO_LOCALIZADO"


def test_document_matched_composition_requires_aprovado_sim_for_name_divergence(tmp_path):
    outro_nome = "PESSOA CADASTRADA DIFERENTE NO ANALITICO"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc=VALID_CPF), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", outro_nome, acquisition="100.00", doc=VALID_CPF)],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    # Aprova a composição e inclui, mas SEM marcar APROVADO=SIM para a
    # divergência de nome (o principal herdou o nome do crédito no preparo,
    # diferente do texto original do PFMI).
    edits = {
        r.values["ID_LINHA"]: {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 1000,
        }
        for r in comp_rows
    }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    principal_row = next(
        r for r in comp_rows if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
    )
    # sugerido do crédito, diverge do PFMI
    assert principal_row.values["NOME_CEDENTE"] == outro_nome
    assert principal_row.values["DOC_CEDENTE"] == VALID_CPF
    assert any("DIVERGENCIA_NOME" in issue for issue in excinfo.value.issues)


def test_document_matched_composition_succeeds_with_aprovado_and_composicao_sim(tmp_path):
    outro_nome = "PESSOA CADASTRADA DIFERENTE NO ANALITICO"
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc=VALID_CPF), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", outro_nome, acquisition="100.00", doc=VALID_CPF)],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {}
    for r in comp_rows:
        is_principal = r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edits[r.values["ID_LINHA"]] = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            # soma exata do nominal_value padrão do crédito (_credit -> 500.00)
            "VL_NOMINAL": 300.00 if is_principal else 200.00,
            "APROVADO": "SIM",  # nome final diverge do PFMI original nos dois casos
            **(
                {}
                if is_principal
                else {
                    "NOME_CEDENTE": "REPRESENTANTE FICTICIO",
                    "DOC_CEDENTE": VALID_CPF,
                    "TIPO_PESSOA_CEDENTE": 1,
                }
            ),
        }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded)
    assert len(validated.rows) == 2
    assert sorted(r["VL_PRESENTE"] for r in validated.rows) == [Decimal("40.00"), Decimal("60.00")]


# 12. Documento inválido bloqueia ---------------------------------------------
def test_invalid_beneficiary_document_blocks_composition():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00", doc="123"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_DOCUMENTO_INVALIDO"


def _write_and_edit(tmp_path, batch, edits):
    path = tmp_path / "intermediario.xlsx"
    write_intermediate(batch, path)
    from openpyxl import load_workbook

    book = load_workbook(path)
    sheet = book["CREDITOS"]
    header = {cell.value: cell.column for cell in sheet[1]}
    for row in range(2, sheet.max_row + 1):
        line_id = sheet.cell(row, header["ID_LINHA"]).value
        if line_id in edits:
            for field, value in edits[line_id].items():
                sheet.cell(row, header[field]).value = value
    book.save(path)
    book.close()
    return path


# 13. Aprovação integral do grupo não bloqueia por inconsistência ------------
def test_full_group_approval_is_internally_consistent(tmp_path):
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00", block="B0001"),
            _pay(REPRESENTANTE, "40.00", block="B0001"),
            _pay(
                "OUTRO CEDENTE FICTICIO", "999.00", block="B0002", failure="FALENCIA FICTICIA DOIS"
            ),
        ],
        [
            _credit("REF-1", PRINCIPAL, acquisition="100.00"),
            _credit(
                "REF-2",
                "OUTRO CEDENTE FICTICIO",
                failure="FALENCIA FICTICIA DOIS",
                acquisition="999.00",
                nominal="1200.00",  # != aquisição: evita base de comissão zero/-0.00
            ),
        ],
        due_records=[_due(), _due("FALENCIA FICTICIA DOIS")],
    )
    comp = batch.compositions[0]
    comp_ids = [r.values["ID_LINHA"] for r in _composition_rows(batch, comp["composicao_id"])]
    other_id = next(
        r.values["ID_LINHA"] for r in batch.rows if r.values["COMPOSICAO_ID"] == ""
    )
    edits = {line_id: {"COMPOSICAO_APROVADA": "SIM"} for line_id in comp_ids}
    edits[other_id] = {"INCLUIR_CNAB": "SIM"}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    # Aprovação consistente (todas as linhas da composição com o mesmo
    # COMPOSICAO_APROVADA e INCLUIR_CNAB=NAO) não deve, por si só, gerar
    # nenhuma mensagem de inconsistência de composição. O lote inteiro ainda
    # falha (regra pré-existente e inalterada: todo grupo precisa de um
    # crédito selecionado antes de qualquer geração), mas por um motivo
    # completamente diferente, já esperado sem composição nenhuma.
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    issues = excinfo.value.issues
    assert not any("COMPOSICAO_APROVADA deve ser igual" in issue for issue in issues)
    assert not any("INCLUIR_CNAB deve ser igual" in issue for issue in issues)
    assert not any("geração de títulos por composição" in issue for issue in issues)
    assert any("nenhum crédito selecionado" in issue for issue in issues)


# 14. Aprovação ou INCLUIR_CNAB parcial bloqueia ------------------------------
def test_partial_composition_approval_is_rejected(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {comp_rows[0].values["ID_LINHA"]: {"COMPOSICAO_APROVADA": "SIM"}}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("COMPOSICAO_APROVADA deve ser igual" in issue for issue in excinfo.value.issues)


def test_partial_inclusion_is_rejected(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {comp_rows[0].values["ID_LINHA"]: {"INCLUIR_CNAB": "SIM"}}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    issues = excinfo.value.issues
    assert any("INCLUIR_CNAB deve ser igual" in issue for issue in issues) or any(
        "só pode ser incluído" in issue for issue in issues
    )


# 15. Geração nunca ocorre parcialmente para uma composição -------------------
def test_inclusion_without_composicao_aprovada_is_rejected(tmp_path):
    # Grupo inteiro com INCLUIR_CNAB=SIM (consistente) mas sem aprovação
    # explícita da composição: continua bloqueado, mesmo sendo "tudo".
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {
        r.values["ID_LINHA"]: {"INCLUIR_CNAB": "SIM", "VL_NOMINAL": 1000}
        for r in comp_rows
    }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("só pode ser incluído" in issue for issue in excinfo.value.issues)


def test_vl_nominal_is_prefilled_with_proportional_suggestion():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="9999.99")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    assert [r.values["VL_NOMINAL"] for r in rows] == [Decimal("5999.99"), Decimal("4000.00")]
    assert all(r.values["DIFERENCA_NOMINAL"] is None for r in rows)
    assert all("PREENCHER_VL_NOMINAL_COMPOSICAO" not in r.values["PENDENCIAS"] for r in rows)


def _fully_approved_edits(comp_rows, nominais):
    edits = {}
    for row, nominal in zip(comp_rows, nominais, strict=True):
        # A satellite's own name never carries the principal's name in
        # parentheses (that annotation only exists in the raw PFMI text for
        # human traceability) - a reviewer types the component's own name.
        nome_pfmi = row.values["NOME_CEDENTE_PFMI"]
        edits[row.values["ID_LINHA"]] = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": nominal,
            "NOME_CEDENTE": nome_pfmi.split(" (", 1)[0],
            "DOC_CEDENTE": VALID_CPF,
            "TIPO_PESSOA_CEDENTE": 1,
        }
    return edits


# 15b. Nominal obrigatório em cada componente ---------------------------------
def test_generation_blocked_when_one_component_nominal_is_missing(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = _fully_approved_edits(comp_rows, [3000, None])
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("VL_NOMINAL" in issue for issue in excinfo.value.issues)


# Geração completa: grupo aprovado, incluído e com nominal válido -----------
def test_generation_succeeds_for_fully_approved_composition_with_valid_nominal(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="5000.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = _fully_approved_edits(comp_rows, [Decimal("3000.00"), Decimal("2000.00")])
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded)
    assert len(validated.rows) == 2
    nominais = sorted(r["VL_NOMINAL"] for r in validated.rows)
    assert nominais == [Decimal("2000.00"), Decimal("3000.00")]
    presentes = sorted(r["VL_PRESENTE"] for r in validated.rows)
    assert presentes == [Decimal("40.00"), Decimal("60.00")]
    sequencias = sorted(int(r["SEU_NUMERO"]) for r in validated.rows)
    assert len(set(sequencias)) == 2  # cada componente com sua própria sequência


def test_generated_txt_has_one_line_per_component_with_own_sequence(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="5000.00")],
        first_sequence=700,
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = _fully_approved_edits(comp_rows, [Decimal("3000.00"), Decimal("2000.00")])
    path = _write_and_edit(tmp_path, batch, edits)
    result = generate_cnab(
        str(path),
        output_path=str(path.parent / "saida.txt"),
        log_directory=str(path.parent / "logs"),
    )
    assert result.detail_count == 2
    assert len(result.compositions) == 1
    resumo = result.compositions[0]
    assert resumo["total_nominal"] == Decimal("5000.00")
    assert sorted(c["vl_nominal"] for c in resumo["componentes"]) == [
        Decimal("2000.00"),
        Decimal("3000.00"),
    ]
    data = Path(result.output_path).read_bytes()
    lines = data.split(b"\r\n")[:-1]
    assert len(lines) == 4  # header + 2 detalhes + trailer


# 16. Sequências separadas e determinísticas ----------------------------------
def test_composition_components_reserve_separate_deterministic_sequences():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
        first_sequence=500,
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    sequences = sorted(r.values["SEU_NUMERO"] for r in rows)
    assert sequences == [500, 501]
    assert len(set(sequences)) == len(sequences)


# 17. Preservação exata dos totais (sem arredondamento) -----------------------
def test_totals_are_preserved_exactly_without_rounding():
    batch = _prepare(
        [_pay(PRINCIPAL, "33.33"), _pay(REPRESENTANTE, "66.67")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="777.77")],
    )
    comp = batch.compositions[0]
    assert comp["total_pfmi"] == Decimal("100.00")
    assert comp["total_analitico"] == Decimal("100.00")
    assert comp["diferenca"] == Decimal("0.00")
    assert comp["nominal_analitico"] == Decimal("777.77")
    rows = _composition_rows(batch, comp["composicao_id"])
    for row in rows:
        assert row.values["COMPOSICAO_TOTAL_PFMI"] == Decimal("100.00")
        assert row.values["COMPOSICAO_NOMINAL_ANALITICO"] == Decimal("777.77")


# 18. Ausência de dupla contagem ----------------------------------------------
def test_principal_alone_is_never_auto_selected_avoiding_double_use():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    principal_row = next(
        r
        for r in batch.rows
        if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL and r.values["ID_CREDITO"]
    )
    assert principal_row.values["INCLUIR_CNAB"] == "NAO"
    # Componente de composição já PROPOSTA não é um crédito individual: não
    # exige seleção manual própria (COMPOSICAO_APROVADA é o gate real), mas
    # oferece nominal sugerido sem aprovar a composição.
    assert "SELECAO_MANUAL_NECESSARIA" not in principal_row.values["PENDENCIAS"]
    assert "PREENCHER_VL_NOMINAL_COMPOSICAO" not in principal_row.values["PENDENCIAS"]


# 19. Escrita e leitura do Excel preserva os campos ---------------------------
def test_excel_round_trip_preserves_composition_fields(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {comp_rows[0].values["ID_LINHA"]: {"COMPOSICAO_APROVADA": "SIM"}}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    reloaded = {r.values["ID_LINHA"]: r.values for r in loaded.rows}
    for line_id in [r.values["ID_LINHA"] for r in comp_rows]:
        assert reloaded[line_id]["COMPOSICAO_ID"] == comp["composicao_id"]
        assert reloaded[line_id]["COMPOSICAO_TOTAL_PFMI"] == Decimal("100.00")
    assert reloaded[comp_rows[0].values["ID_LINHA"]]["COMPOSICAO_APROVADA"] == "SIM"


# 20. Alteração indevida de campo protegido invalida o manifesto -------------
def test_tampering_with_protected_composition_field_is_rejected(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    comp_rows = _composition_rows(batch, comp["composicao_id"])
    edits = {comp_rows[0].values["ID_LINHA"]: {"COMPOSICAO_TOTAL_PFMI": 12345.67}}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    assert any(
        "COMPOSICAO_TOTAL_PFMI" in issue and "alterad" in issue
        for issue in loaded.structural_issues
    )


# 21. Apresentação das informações na interface -------------------------------
def test_gui_composition_message_is_transparent_and_never_asserts_identity():
    comp = {
        "participantes": f"{PRINCIPAL}; {REPRESENTANTE}",
        "total_pfmi": Decimal("100.00"),
        "total_analitico": Decimal("100.00"),
        "diferenca": Decimal("0.00"),
        "nominal_analitico": Decimal("500.00"),
        "motivo": "Soma exata dos pagamentos do PFMI confere com a aquisição esperada do crédito",
        "estado": "PROPOSTA",
    }
    message = _composition_message(comp)
    assert "Composição encontrada — confira antes de aprovar" in message
    assert "não confirma identidade" in message
    assert "COMPOSICAO_APROVADA" in message
    assert PRINCIPAL in message and REPRESENTANTE in message


def test_gui_blocked_composition_message_says_blocked():
    comp = {
        "participantes": PRINCIPAL,
        "total_pfmi": Decimal("100.01"),
        "total_analitico": Decimal("100.00"),
        "diferenca": Decimal("0.01"),
        "nominal_analitico": None,
        "motivo": "Diferença de +0.01",
        "estado": "BLOQUEADA_DIVERGENCIA_VALOR",
    }
    message = _composition_message(comp)
    assert "Composição bloqueada" in message
    assert "indisponível" in message  # nominal ausente não é inventado


# 22. Preservação dos casos ordinários existentes -----------------------------
def test_ordinary_single_group_case_is_unaffected():
    batch = _prepare(
        [_pay(PRINCIPAL, "100.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    assert batch.compositions == []
    assert batch.rows[0].values["INCLUIR_CNAB"] == "SIM"
    assert batch.rows[0].values["COMPOSICAO_ID"] == ""


# 23. Preservação de CONEX e da comissão --------------------------------------
def test_conexcred_and_commission_are_unaffected_by_composition_feature():
    from gerador_cnab_nox.models import CESSAO

    payment = PfmiPayment(
        block_id="B0001",
        source_row=1,
        title="FUNDO",
        failure=FALENCIA,
        cedent_name=PRINCIPAL,
        beneficiary_name="CONEXCRED INTERMEDIACAO",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100.00"),
        beneficiary_value=Decimal("100.00"),
        modality=CESSAO,
        commission_text="15",
    )
    credit = _credit("REF-1", PRINCIPAL, acquisition="85.00", nominal="100.00")
    batch = _prepare([payment], [credit])
    row = batch.rows[0].values
    assert row["NOME_CEDENTE"] == "CONEXCRED INTERMEDIACAO"
    assert row["COMISSAO_CALCULADA"] == Decimal("2.25")  # 15% sobre base de 15.00
    assert row["COMPOSICAO_ID"] == ""


# 24. Preservação das proteções atuais de eligible ----------------------------
def test_multiple_principal_credits_still_require_manual_selection():
    batch = _prepare(
        [_pay(PRINCIPAL, "40.00", block="B0001"), _pay(PRINCIPAL, "60.00", block="B0002")],
        [
            _credit("REF-A", PRINCIPAL, acquisition="40.00"),
            _credit("REF-B", PRINCIPAL, acquisition="60.00"),
        ],
    )
    rows = [r for r in batch.rows if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL]
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in rows)
    assert batch.compositions == []


# =============================================================================
# Fallback por documento no matching comum (não composição) e seleção manual
# =============================================================================

SOLO_PFMI = "SOLO FICTICIO PAGADOR"
SOLO_ANALITICO = "SOLO FICTICIO CADASTRADO DIFERENTE"
OUTRA_FALENCIA = "FALENCIA FICTICIA DIVERGENTE"


def _write_analytic_xlsx(path, rows):
    wb = Workbook()
    wb.active.append(
        ["Cpf", "Contato", "Campanha", "Valor Receber", "Valor Aquisicao", "Data Assinatura"]
    )
    for row in rows:
        wb.active.append(row)
    wb.save(path)
    wb.close()


def test_ordinary_group_matches_by_document_when_name_diverges_same_failure():
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.00", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_ANALITICO, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0].values
    assert row["ID_CREDITO"]  # achou o crédito por documento
    assert row["INCLUIR_CNAB"] == "NAO"  # nunca automático
    assert "SELECAO_MANUAL_NECESSARIA" in row["PENDENCIAS"]
    assert "DIVERGENCIA_NOME" in row["PENDENCIAS"]
    assert "CREDOR_LOCALIZADO_POR_DOCUMENTO" in row["PENDENCIAS"]


def test_ordinary_document_fallback_never_crosses_failure():
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.00", doc=VALID_CPF)],
        [
            _credit(
                "REF-1", SOLO_ANALITICO, failure=OUTRA_FALENCIA, acquisition="100.00", doc=VALID_CPF
            )
        ],
    )
    row = batch.rows[0].values
    assert row["ID_CREDITO"] == ""
    assert "CREDITO_NAO_LOCALIZADO" in row["PENDENCIAS"]


def test_ordinary_ambiguous_document_is_never_an_automatic_candidate():
    # Aquisições diferentes garantem que nenhum dos dois bata por valor (regra
    # já existente, sem relação com esta funcionalidade); assim isola-se
    # exclusivamente o comportamento do fallback por documento.
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.00", doc=VALID_CPF)],
        [
            _credit("REF-A", "OUTRO CADASTRO A", acquisition="41.00", doc=VALID_CPF),
            _credit("REF-B", "OUTRO CADASTRO B", acquisition="42.00", doc=VALID_CPF),
        ],
    )
    row = batch.rows[0].values
    assert row["ID_CREDITO"] == ""
    assert "CREDITO_NAO_LOCALIZADO" in row["PENDENCIAS"]


def _cross_failure_batch():
    # Documento do PFMI não corresponde a nenhum crédito e a falência também
    # diverge: matching automático (nome ou documento) não resolve nada.
    return _prepare(
        [_pay(SOLO_PFMI, "100.00", doc="11111111111")],
        [
            _credit(
                "REF-1", SOLO_ANALITICO, failure=OUTRA_FALENCIA,
                acquisition="100.00", nominal="777.00", doc=VALID_CPF,
            )
        ],
    )


def _manual_credit(reference="REF-1", cedent_name=SOLO_ANALITICO, present="100.00", source_row=2):
    return AnalyticCredit(
        source_row=source_row,
        reference=reference,
        cedent_document=VALID_CPF,
        cedent_name=cedent_name,
        campaign=OUTRA_FALENCIA,
        nominal_value=Decimal("777.00"),
        acquisition_value=Decimal(present),
        signature_date=date(2026, 9, 1),
    )


def _base_manual_edits(row, **overrides):
    edits = {
        "SELECAO_MANUAL_DOCUMENTO": VALID_CPF,
        "SELECAO_MANUAL_APROVADA": "SIM",
        "APROVADO": "SIM",
        "NOME_CEDENTE": SOLO_ANALITICO,
        "DOC_CEDENTE": VALID_CPF,
        "TIPO_PESSOA_CEDENTE": 1,
        "VL_NOMINAL": 777.00,
        "VL_PRESENTE": 100.00,  # próprio total do PFMI (regra já existente)
        "DT_EMISSAO_TITULO": date(2026, 9, 1),
        "INCLUIR_CNAB": "SIM",
    }
    edits.update(overrides)
    return {row.values["ID_LINHA"]: edits}


def test_manual_selection_resolves_cross_failure_case_with_full_approval(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    assert row.values["ID_CREDITO"] == ""  # confirma que nada automático achou
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded, analytic=[_manual_credit()])
    assert len(validated.rows) == 1
    assert validated.rows[0]["VL_NOMINAL"] == Decimal("777.00")
    assert validated.rows[0]["VL_PRESENTE"] == Decimal("100.00")


def test_manual_selection_blocked_without_selecao_manual_aprovada(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row, SELECAO_MANUAL_APROVADA="")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[_manual_credit()])
    assert any("SELECAO_MANUAL_APROVADA" in issue for issue in excinfo.value.issues)


def test_manual_selection_blocked_without_aprovado_for_name_divergence(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row, APROVADO="")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[_manual_credit()])
    assert any("DIVERGENCIA_NOME" in issue for issue in excinfo.value.issues)


def test_manual_selection_blocked_when_document_not_found(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[])  # nenhum crédito com esse documento
    assert any("nenhum crédito" in issue for issue in excinfo.value.issues)


def test_manual_selection_blocked_when_document_ambiguous(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    duplicate = _manual_credit(reference="REF-2", cedent_name="OUTRO NOME", source_row=3)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[_manual_credit(), duplicate])
    assert any("mais de um crédito" in issue for issue in excinfo.value.issues)


def test_manual_selection_requires_exact_nominal_no_rounding(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row, VL_NOMINAL=777.01)  # um centavo a mais
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[_manual_credit()])
    assert any("VL_NOMINAL" in issue for issue in excinfo.value.issues)


def test_manual_selection_requires_exact_financial_closure(tmp_path):
    # Crédito cuja aquisição não fecha com o total do PFMI (100.00): bloqueia,
    # mesmo com documento exato e todas as aprovações presentes.
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    off_by_a_cent = _manual_credit(present="100.01")
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[off_by_a_cent])
    assert any("não fecha exatamente" in issue for issue in excinfo.value.issues)


def test_manual_selection_end_to_end_generates_real_txt(tmp_path):
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    analytic_path = tmp_path / "analitico_ficticio.xlsx"
    _write_analytic_xlsx(
        analytic_path,
        [[VALID_CPF, SOLO_ANALITICO, OUTRA_FALENCIA, 777.00, 100.00, "01/09/2026"]],
    )
    result = generate_cnab(
        str(path),
        output_path=str(tmp_path / "saida_manual.txt"),
        log_directory=str(tmp_path / "logs"),
        analytic_path=str(analytic_path),
    )
    assert result.detail_count == 1
    assert result.total_nominal == Decimal("777.00")
    assert result.total_present == Decimal("100.00")


def test_manual_selection_without_analytic_path_stays_blocked(tmp_path):
    # Sem o Analítico disponível na geração, a seleção manual não pode ser
    # conferida - nunca aceita "de qualquer jeito".
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row)
    path = _write_and_edit(tmp_path, batch, edits)
    with pytest.raises(ValidationError) as excinfo:
        generate_cnab(
            str(path),
            output_path=str(tmp_path / "nao_deve_existir.txt"),
            log_directory=str(tmp_path / "logs"),
        )
    assert any("nenhum crédito" in issue for issue in excinfo.value.issues)


# Seleção manual do PRINCIPAL de uma composição quando a detecção automática
# falha (nome e documento não localizam nenhum crédito na mesma falência) ----
# Identidade por REGISTRO ÚNICO (aba + linha física do Analítico), nunca por
# um campo de campanha/lote compartilhado (ex.: "Prospect" no arquivo real,
# que se repete em centenas de créditos - ver check_teste07_doc_match.py).

SEM_CREDITO_PRINCIPAL = "PRINCIPAL SEM CREDITO FICTICIO"
SEM_CREDITO_SATELITE = f"REPRESENTANTE SEM CREDITO ({SEM_CREDITO_PRINCIPAL})"
CREDITO_ANALITICO_DIVERGENTE = "CREDOR ANALITICO NOME DIVERGENTE FICTICIO"
LOCALIZACAO_ABA = "Sheet1"
LOCALIZACAO_LINHA = 10


def _unlocatable_composition_batch(acquisition="100.00"):
    return _prepare(
        [_pay(SEM_CREDITO_PRINCIPAL, "60.00"), _pay(SEM_CREDITO_SATELITE, "40.00")],
        [
            _credit(
                "2100.0",  # campo de campanha/lote fictício, deliberadamente ignorado
                CREDITO_ANALITICO_DIVERGENTE,
                acquisition=acquisition,
                doc=VALID_CNPJ,
                source_sheet=LOCALIZACAO_ABA,
                source_row=LOCALIZACAO_LINHA,
            )
        ],
    )


def _composition_manual_edits(rows, aba, linha, **overrides):
    edits = {}
    for row, nominal in zip(rows, ("300.00", "200.00"), strict=False):
        edit = {
            "COMPOSICAO_SELECAO_MANUAL_ABA": aba,
            "COMPOSICAO_SELECAO_MANUAL_LINHA": linha,
            "COMPOSICAO_APROVADA": "SIM",
            "APROVADO": "SIM",
            "NOME_CEDENTE": CREDITO_ANALITICO_DIVERGENTE,
            "DOC_CEDENTE": VALID_CNPJ,
            "TIPO_PESSOA_CEDENTE": 2,
            "VL_NOMINAL": float(nominal),
            "DT_EMISSAO_TITULO": date(2026, 9, 1),
            "INCLUIR_CNAB": "SIM",
        }
        edit.update(overrides)
        edits[row.values["ID_LINHA"]] = edit
    return edits


def test_composition_blocked_without_automatic_match_has_no_credit():
    batch = _unlocatable_composition_batch()
    assert len(batch.compositions) == 1
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_CREDITO_NAO_LOCALIZADO"
    rows = _composition_rows(batch, comp["composicao_id"])
    assert len(rows) == 2
    assert all(r.values["ID_CREDITO"] == "" for r in rows)


def test_composition_candidates_are_offered_for_human_review():
    # matching.py::_composition_manual_candidates: nome, documento mascarado
    # (nunca o documento completo), valor e localizacao - nunca seleciona
    # sozinho, so aparece para conferencia.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    candidatos = json.loads(rows[0].values["COMPOSICAO_CANDIDATOS"])
    assert len(candidatos) == 1
    candidato = candidatos[0]
    assert candidato["aba"] == LOCALIZACAO_ABA
    assert candidato["linha"] == LOCALIZACAO_LINHA
    assert candidato["nome"] == CREDITO_ANALITICO_DIVERGENTE
    assert VALID_CNPJ not in candidato["documento_mascarado"]
    assert candidato["documento_mascarado"].count("*") > 0


def test_composition_manual_location_resolves_when_sum_closes_exactly(tmp_path):
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    validated = validate_for_generation(loaded, analytic=[credit])
    assert len(validated.rows) == 2
    assert sum((r["VL_PRESENTE"] for r in validated.rows), Decimal("0.00")) == Decimal("100.00")


def test_composition_manual_location_end_to_end_generates_real_txt(tmp_path):
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ
    )
    analytic_path = tmp_path / "analitico_composicao.xlsx"
    _write_analytic_xlsx(
        analytic_path,
        [
            (
                credit.cedent_document,
                credit.cedent_name,
                credit.campaign,
                float(credit.nominal_value),
                float(credit.acquisition_value),
                credit.signature_date,
            )
        ],
    )
    # openpyxl nomeia a primeira aba de um Workbook novo como "Sheet"; o
    # único registro de dados fica na linha física 2 (linha 1 é o cabeçalho)
    # - exatamente o que o operador leria conferindo o Analítico real.
    edits = _composition_manual_edits(rows, "Sheet", 2)
    path = _write_and_edit(tmp_path, batch, edits)
    txt_out = tmp_path / "saida.txt"
    result = generate_cnab(
        str(path),
        output_path=str(txt_out),
        log_directory=str(tmp_path / "logs"),
        analytic_path=str(analytic_path),
    )
    assert result.detail_count == 2
    assert txt_out.exists()


def test_composition_manual_location_inconsistent_across_rows_stays_blocked(tmp_path):
    # Adulteração/erro: o satélite aponta para uma linha diferente da do
    # principal. Auditoria exige que todo componente aponte para o mesmo
    # registro, nunca um "atalho" por uma única linha.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA)
    edits[rows[1].values["ID_LINHA"]]["COMPOSICAO_SELECAO_MANUAL_LINHA"] = LOCALIZACAO_LINHA + 1
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("devem ser preenchidos e iguais" in issue for issue in excinfo.value.issues)


def test_composition_manual_location_only_aba_filled_stays_blocked(tmp_path):
    # Adulteração/erro: preencheu só a aba, sem a linha (ou vice-versa) - os
    # dois campos são exigidos juntos, nunca um identificador parcial.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, "")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("informe COMPOSICAO_SELECAO_MANUAL_ABA" in issue for issue in excinfo.value.issues)


def test_composition_manual_location_tampered_nonexistent_row_stays_blocked(tmp_path):
    # Adulteração: linha digitada não existe de fato no Analítico carregado.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA + 999)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any(
        "nenhum crédito do Analítico foi encontrado" in issue for issue in excinfo.value.issues
    )


def test_composition_manual_location_tampered_wrong_sheet_stays_blocked(tmp_path):
    # Adulteração: linha certa, aba errada - não é a mesma coordenada.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, "OutraAba", LOCALIZACAO_LINHA)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any(
        "nenhum crédito do Analítico foi encontrado" in issue for issue in excinfo.value.issues
    )


def test_composition_manual_location_ambiguous_duplicate_row_stays_blocked(tmp_path):
    # Ambiguidade defensiva: dois créditos apontam para a mesma coordenada
    # (aba, linha) no Analítico carregado - nunca escolhe "o primeiro".
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    duplicate = _credit(
        "2100.0", "OUTRO CREDOR MESMA LINHA", acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit, duplicate])
    assert any("mais de um crédito" in issue for issue in excinfo.value.issues)


def test_composition_manual_location_reused_by_two_compositions_stays_blocked(tmp_path):
    # Adulteração: a mesma linha do Analítico é usada para "resolver" duas
    # composições diferentes - o mesmo crédito não pode ser reaproveitado
    # fora da própria composição (mecanismo já existente de dedup).
    outro_principal = "OUTRO PRINCIPAL SEM CREDITO FICTICIO"
    outro_satelite = f"OUTRO REPRESENTANTE SEM CREDITO ({outro_principal})"
    batch = _prepare(
        [
            _pay(SEM_CREDITO_PRINCIPAL, "60.00", block="B0001"),
            _pay(SEM_CREDITO_SATELITE, "40.00", block="B0001"),
            _pay(outro_principal, "60.00", block="B0002"),
            _pay(outro_satelite, "40.00", block="B0002"),
        ],
        [
            _credit(
                "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
                source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
            )
        ],
    )
    assert len(batch.compositions) == 2
    edits = {}
    for comp in batch.compositions:
        rows = _composition_rows(batch, comp["composicao_id"])
        edits.update(_composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA))
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any(
        "o mesmo crédito do Analítico foi selecionado mais de uma vez" in issue
        for issue in excinfo.value.issues
    )


def test_composition_manual_location_one_cent_difference_stays_blocked(tmp_path):
    # Soma dos componentes = 100.00; aquisicao do credito = 100.01: mesmo um
    # centavo de diferenca bloqueia, sem arredondamento nem tolerancia.
    batch = _unlocatable_composition_batch(acquisition="100.01")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.01", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("não fecha exatamente" in issue for issue in excinfo.value.issues)


def test_composition_approval_without_manual_location_still_blocked(tmp_path):
    # Sem nenhuma tentativa de selecao manual (campos em branco em todas as
    # linhas), o comportamento de hoje se mantem: COMPOSICAO_APROVADA=SIM
    # continua proibido enquanto o estado nao for PROPOSTA.
    batch = _unlocatable_composition_batch(acquisition="100.00")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, "", "")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.00", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("só é permitido quando" in issue for issue in excinfo.value.issues)


# Aprovação manual auditável de divergência de VALOR entre PFMI e Analítico:
# quando o crédito já está identificado sem ambiguidade (nome, documento,
# localização de composição) mas o Valor Aquisição do Analítico não fecha
# exatamente com o total do PFMI, a Lu pode aprovar explicitamente o uso do
# valor do PFMI - nunca automático, nunca arredondado, sempre auditável -----

DIVERGENCIA_JUSTIFICATIVA = "Instrumento conferido; usar o valor exato da PFMI."


def _divergencia_edits(row, **overrides):
    edits = {
        "INCLUIR_CNAB": "SIM",
        "DIVERGENCIA_VALOR_APROVADA": "SIM",
        "DIVERGENCIA_VALOR_JUSTIFICATIVA": DIVERGENCIA_JUSTIFICATIVA,
    }
    edits.update(overrides)
    return {row.values["ID_LINHA"]: edits}


def test_ordinary_value_divergence_fields_are_populated_at_prepare_time():
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0].values
    assert row["ID_CREDITO"]  # identidade já confirmada por nome, sem ambiguidade
    assert row["DIVERGENCIA_VALOR_PFMI"] == Decimal("100.01")
    assert row["DIVERGENCIA_VALOR_ANALITICO"] == Decimal("100.00")
    assert row["DIVERGENCIA_VALOR_DIFERENCA"] == Decimal("0.01")
    assert row["DIVERGENCIA_VALOR_APROVADA"] == ""
    assert row["DIVERGENCIA_VALOR_JUSTIFICATIVA"] == ""


def test_case_without_divergence_has_blank_audit_fields_and_needs_no_approval(tmp_path):
    # Casos sem divergência continuam funcionando sem aprovação nenhuma.
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.00", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    assert row.values["DIVERGENCIA_VALOR_PFMI"] is None
    assert row.values["DIVERGENCIA_VALOR_ANALITICO"] is None
    assert row.values["DIVERGENCIA_VALOR_DIFERENCA"] is None
    edits = {row.values["ID_LINHA"]: {"INCLUIR_CNAB": "SIM"}}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(
        loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
    )
    assert len(validated.rows) == 1
    assert validated.rows[0]["VL_PRESENTE"] == Decimal("100.00")


def test_ordinary_value_divergence_blocked_without_approval(tmp_path):
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    path = _write_and_edit(tmp_path, batch, {row.values["ID_LINHA"]: {"INCLUIR_CNAB": "SIM"}})
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("DIVERGENCIA_VALOR" in issue for issue in excinfo.value.issues)


def test_ordinary_value_divergence_blocked_without_justificativa(tmp_path):
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    edits = _divergencia_edits(row, DIVERGENCIA_VALOR_JUSTIFICATIVA="")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("JUSTIFICATIVA é obrigatória" in issue for issue in excinfo.value.issues)


def test_ordinary_justificativa_without_approval_is_blocked(tmp_path):
    # Adulteração: justificativa preenchida mas sem aprovação explícita.
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    edits = {
        row.values["ID_LINHA"]: {
            "INCLUIR_CNAB": "SIM",
            "DIVERGENCIA_VALOR_JUSTIFICATIVA": DIVERGENCIA_JUSTIFICATIVA,
        }
    }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("só é aceita quando" in issue for issue in excinfo.value.issues)


def test_ordinary_value_divergence_invalid_approval_value_is_blocked(tmp_path):
    # Adulteração: qualquer valor diferente de SIM/vazio é rejeitado (não
    # existe "NAO" explícito para esta aprovação).
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    edits = _divergencia_edits(row, DIVERGENCIA_VALOR_APROVADA="TALVEZ")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("DIVERGENCIA_VALOR_APROVADA deve ser SIM" in issue for issue in excinfo.value.issues)


def test_ordinary_value_divergence_approved_uses_pfmi_value_positive_difference(tmp_path):
    # PFMI > Analítico (diferença positiva).
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    path = _write_and_edit(tmp_path, batch, _divergencia_edits(row))
    loaded = read_intermediate(path)
    validated = validate_for_generation(
        loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
    )
    assert len(validated.rows) == 1
    assert validated.rows[0]["VL_PRESENTE"] == Decimal("100.01")


def test_ordinary_value_divergence_approved_uses_pfmi_value_negative_difference(tmp_path):
    # PFMI < Analítico (diferença negativa).
    batch = _prepare(
        [_pay(SOLO_PFMI, "99.98", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    path = _write_and_edit(tmp_path, batch, _divergencia_edits(row))
    loaded = read_intermediate(path)
    validated = validate_for_generation(
        loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
    )
    assert len(validated.rows) == 1
    assert validated.rows[0]["VL_PRESENTE"] == Decimal("99.98")


def test_ordinary_value_divergence_generates_txt_with_exact_pfmi_value(tmp_path):
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    path = _write_and_edit(tmp_path, batch, _divergencia_edits(row))
    txt_out = tmp_path / "saida.txt"
    result = generate_cnab(
        str(path), output_path=str(txt_out), log_directory=str(tmp_path / "logs")
    )
    assert result.total_present == Decimal("100.01")
    data = txt_out.read_bytes()
    detail = data.split(b"\r\n")[1]
    assert Decimal(detail[192:205].decode()) / 100 == Decimal("100.01")


def test_ordinary_credit_absent_cannot_be_approved_by_value_alone(tmp_path):
    # Crédito ausente: aprovar a divergência de valor nunca inventa uma
    # identidade nem dispensa a seleção manual por documento.
    batch = _cross_failure_batch()
    row = batch.rows[0]
    assert row.values["ID_CREDITO"] == ""
    edits = _divergencia_edits(row, INCLUIR_CNAB="SIM")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[_manual_credit()])
    assert any(
        "linha sem crédito original do Analítico não pode ser selecionada" in issue
        for issue in excinfo.value.issues
    )


def test_ordinary_manual_document_value_divergence_blocked_without_approval(tmp_path):
    # Identidade confirmada por documento (seleção manual, cruzando
    # falência), mas o valor diverge - continua bloqueado sem aprovação.
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(row, VL_NOMINAL=777.00)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _manual_credit(present="100.01")  # Analítico diverge 1 centavo do PFMI (100.00)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("não fecha exatamente" in issue for issue in excinfo.value.issues)


def test_ordinary_manual_document_value_divergence_approved_successfully(tmp_path):
    # Caso real do Teste 07 (Wilson): identidade só por documento manual,
    # cruzando falência, e o valor também diverge - as duas aprovações são
    # independentes e as duas são exigidas.
    batch = _cross_failure_batch()
    row = batch.rows[0]
    edits = _base_manual_edits(
        row,
        VL_NOMINAL=777.00,
        DIVERGENCIA_VALOR_APROVADA="SIM",
        DIVERGENCIA_VALOR_JUSTIFICATIVA=DIVERGENCIA_JUSTIFICATIVA,
    )
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _manual_credit(present="100.01")
    validated = validate_for_generation(loaded, analytic=[credit])
    assert len(validated.rows) == 1
    assert validated.rows[0]["VL_PRESENTE"] == Decimal("100.00")  # valor exato do PFMI


def test_composition_value_divergence_fields_are_populated_and_credit_kept():
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "39.99")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_DIVERGENCIA_VALOR"
    assert comp["diferenca"] == Decimal("-0.01")
    rows = _composition_rows(batch, comp["composicao_id"])
    assert len(rows) == 2
    # Identidade do principal já confirmada mesmo com o valor divergente:
    # nome/documento/nominal sugeridos ficam disponíveis para conferência.
    assert all(r.values["ID_CREDITO"] for r in rows)
    assert all(r.values["DIVERGENCIA_VALOR_PFMI"] == Decimal("99.99") for r in rows)
    assert all(r.values["DIVERGENCIA_VALOR_ANALITICO"] == Decimal("100.00") for r in rows)
    assert all(r.values["DIVERGENCIA_VALOR_DIFERENCA"] == Decimal("-0.01") for r in rows)


def _composition_divergence_edits(rows, **overrides):
    edits = {}
    for r in rows:
        is_principal = r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edit = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "APROVADO": "SIM",
            "DIVERGENCIA_VALOR_APROVADA": "SIM",
            "DIVERGENCIA_VALOR_JUSTIFICATIVA": DIVERGENCIA_JUSTIFICATIVA,
            "VL_NOMINAL": 1000,
        }
        if not is_principal:
            edit.update(
                {"NOME_CEDENTE": "REPRESENTANTE FICTICIO", "DOC_CEDENTE": VALID_CPF,
                 "TIPO_PESSOA_CEDENTE": 1}
            )
        edit.update(overrides)
        edits[r.values["ID_LINHA"]] = edit
    return edits


def test_composition_value_divergence_blocked_without_approval(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "39.99")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_divergence_edits(rows, DIVERGENCIA_VALOR_APROVADA="")
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("só é permitido quando" in issue for issue in excinfo.value.issues)


def test_composition_value_divergence_inconsistent_across_rows_is_blocked(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "39.99")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_divergence_edits(rows)
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] != PRINCIPAL)
    edits[satellite.values["ID_LINHA"]]["DIVERGENCIA_VALOR_APROVADA"] = ""
    edits[satellite.values["ID_LINHA"]]["DIVERGENCIA_VALOR_JUSTIFICATIVA"] = ""
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    issues = excinfo.value.issues
    assert any("DIVERGENCIA_VALOR_APROVADA deve ser igual" in issue for issue in issues) or any(
        "só é permitido quando" in issue for issue in issues
    )


def test_composition_value_divergence_approved_successfully_negative_difference(tmp_path):
    # nominal="2000.00": soma exata dos VL_NOMINAL de _composition_divergence_edits
    # (1000 + 1000), já que a nova checagem de soma nominal (regra 5) exige
    # que feche exatamente com o nominal do crédito.
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "39.99")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="2000.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_divergence_edits(rows)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded)
    assert len(validated.rows) == 2
    assert sum((r["VL_PRESENTE"] for r in validated.rows), Decimal("0.00")) == Decimal("99.99")


def test_composition_value_divergence_approved_successfully_positive_difference(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.01")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="2000.00")],
    )
    comp = batch.compositions[0]
    assert comp["diferenca"] == Decimal("0.01")
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_divergence_edits(rows)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded)
    assert len(validated.rows) == 2
    assert sum((r["VL_PRESENTE"] for r in validated.rows), Decimal("0.00")) == Decimal("100.01")


def test_composition_value_divergence_approval_never_resolves_ambiguous_credit(tmp_path):
    # Ambiguidade de crédito (dois créditos possíveis para o principal):
    # aprovar a divergência de valor nunca resolve identidade/ambiguidade.
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [
            _credit("REF-1", PRINCIPAL, acquisition="100.01"),
            _credit("REF-2", PRINCIPAL, acquisition="100.01"),
        ],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_MULTIPLOS_CREDITOS"
    all_rows = _composition_rows(batch, comp["composicao_id"])
    # O principal ambíguo gera uma linha candidata por crédito possível (regra
    # já existente e comum a qualquer matching ambíguo); escolhe uma só, como
    # a Lu faria ao marcar INCLUIR_CNAB=SIM numa única linha física.
    seen_groups: set[str] = set()
    rows = []
    for r in all_rows:
        if r.values["ID_GRUPO"] in seen_groups:
            continue
        seen_groups.add(r.values["ID_GRUPO"])
        rows.append(r)
    edits = _composition_divergence_edits(rows)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded,
            analytic=[
                _credit("REF-1", PRINCIPAL, acquisition="100.01"),
                _credit("REF-2", PRINCIPAL, acquisition="100.01"),
            ],
        )
    assert any("só é permitido quando" in issue for issue in excinfo.value.issues)


def test_divergencia_valor_readonly_fields_cannot_be_tampered(tmp_path):
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    edits = _divergencia_edits(row, DIVERGENCIA_VALOR_ANALITICO=999.99)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("proveniência" in issue and "alterado" in issue for issue in excinfo.value.issues)


def test_divergencia_valor_pfmi_readonly_field_cannot_be_tampered(tmp_path):
    batch = _prepare(
        [_pay(SOLO_PFMI, "100.01", doc=VALID_CPF)],
        [_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0]
    edits = _divergencia_edits(row, DIVERGENCIA_VALOR_PFMI=1.00)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(
            loaded, analytic=[_credit("REF-1", SOLO_PFMI, acquisition="100.00", doc=VALID_CPF)]
        )
    assert any("proveniência" in issue and "alterado" in issue for issue in excinfo.value.issues)


# Caso real do Teste 07 (Gobbo 1 / Valmiro): o principal da composição não é
# encontrado automaticamente (precisa da localização manual por aba+linha,
# V1-4) E o valor também diverge - as duas aprovações são independentes.


def test_composition_manual_location_value_divergence_blocked_without_approval(tmp_path):
    batch = _unlocatable_composition_batch(acquisition="100.01")
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_CREDITO_NAO_LOCALIZADO"
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(rows, LOCALIZACAO_ABA, LOCALIZACAO_LINHA)
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.01", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded, analytic=[credit])
    assert any("não fecha exatamente" in issue for issue in excinfo.value.issues)


def test_composition_manual_location_value_divergence_approved_successfully(tmp_path):
    batch = _unlocatable_composition_batch(acquisition="100.01")
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = _composition_manual_edits(
        rows,
        LOCALIZACAO_ABA,
        LOCALIZACAO_LINHA,
        DIVERGENCIA_VALOR_APROVADA="SIM",
        DIVERGENCIA_VALOR_JUSTIFICATIVA=DIVERGENCIA_JUSTIFICATIVA,
    )
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    credit = _credit(
        "2100.0", CREDITO_ANALITICO_DIVERGENTE, acquisition="100.01", doc=VALID_CNPJ,
        source_sheet=LOCALIZACAO_ABA, source_row=LOCALIZACAO_LINHA,
    )
    validated = validate_for_generation(loaded, analytic=[credit])
    assert len(validated.rows) == 2
    assert sum((r["VL_PRESENTE"] for r in validated.rows), Decimal("0.00")) == Decimal("100.00")


# Regressão: composição PROPOSTA (soma exata) não deve gerar pendências de
# crédito individual (DIVERGENCIA_VALOR no principal, SELECAO_MANUAL_NECESSARIA
# no satélite) - o gate real é COMPOSICAO_APROVADA. Duas composições no mesmo
# lote (uma com 2 componentes, outra com 3) para garantir que cada uma é
# tratada de forma independente e ambas geram corretamente.

OUTRO_PRINCIPAL_DOIS = "OUTRO PRINCIPAL FICTICIO DOIS"
OUTRO_REPRESENTANTE_DOIS = f"OUTRO REPRESENTANTE FICTICIO ({OUTRO_PRINCIPAL_DOIS})"


def test_two_and_three_component_compositions_in_same_batch_have_no_individual_pendencies():
    batch = _prepare(
        [
            # Composição de 3 componentes.
            _pay(PRINCIPAL, "40.00", block="B0001"),
            _pay(REPRESENTANTE, "20.00", block="B0001"),
            _pay(ESCRITORIO, "10.00", block="B0001"),
            # Composição de 2 componentes, credor e crédito independentes.
            _pay(OUTRO_PRINCIPAL_DOIS, "60.00", block="B0002"),
            _pay(OUTRO_REPRESENTANTE_DOIS, "40.00", block="B0002"),
        ],
        [
            _credit("REF-1", PRINCIPAL, acquisition="70.00"),
            _credit("REF-2", OUTRO_PRINCIPAL_DOIS, acquisition="100.00"),
        ],
    )
    assert len(batch.compositions) == 2
    comp_tres = next(c for c in batch.compositions if len(c["component_ids"]) == 3)
    comp_dois = next(c for c in batch.compositions if len(c["component_ids"]) == 2)
    assert comp_tres["estado"] == "PROPOSTA"
    assert comp_dois["estado"] == "PROPOSTA"

    for comp in (comp_tres, comp_dois):
        rows = _composition_rows(batch, comp["composicao_id"])
        principal_name = comp["participantes"].split(";")[0].strip()
        for row in rows:
            v = row.values
            assert "SELECAO_MANUAL_NECESSARIA" not in v["PENDENCIAS"]
            assert "CREDITO_NAO_LOCALIZADO" not in v["PENDENCIAS"]
            if v["NOME_CEDENTE_PFMI"] == principal_name:
                assert "DIVERGENCIA_VALOR" not in v["PENDENCIAS"]
            assert "PREENCHER_VL_NOMINAL_COMPOSICAO" not in v["PENDENCIAS"]
            assert v["INCLUIR_CNAB"] == "NAO"  # nunca automático


def test_two_and_three_component_compositions_generate_successfully(tmp_path):
    batch = _prepare(
        [
            _pay(PRINCIPAL, "40.00", block="B0001"),
            _pay(REPRESENTANTE, "20.00", block="B0001"),
            _pay(ESCRITORIO, "10.00", block="B0001"),
            _pay(OUTRO_PRINCIPAL_DOIS, "60.00", block="B0002"),
            _pay(OUTRO_REPRESENTANTE_DOIS, "40.00", block="B0002"),
        ],
        [
            # nominal = 500.00 x quantidade de componentes: soma exata dos
            # VL_NOMINAL abaixo, já que a checagem de soma nominal (regra 5)
            # exige que feche exatamente com o nominal do crédito.
            _credit("REF-1", PRINCIPAL, acquisition="70.00", nominal="1500.00"),
            _credit("REF-2", OUTRO_PRINCIPAL_DOIS, acquisition="100.00", nominal="1000.00"),
        ],
    )
    edits = {}
    for comp in batch.compositions:
        rows = _composition_rows(batch, comp["composicao_id"])
        principal_name = comp["participantes"].split(";")[0].strip()
        for row in rows:
            is_principal = row.values["NOME_CEDENTE_PFMI"] == principal_name
            edit = {
                "COMPOSICAO_APROVADA": "SIM",
                "INCLUIR_CNAB": "SIM",
                "VL_NOMINAL": 500.00,
            }
            if not is_principal:
                # satélite nunca é pré-preenchido a partir do crédito: precisa
                # de nome/documento próprios e aprovação da divergência de nome.
                edit.update(
                    {
                        "APROVADO": "SIM",
                        "NOME_CEDENTE": "SATELITE FICTICIO",
                        "DOC_CEDENTE": VALID_CPF,
                        "TIPO_PESSOA_CEDENTE": 1,
                    }
                )
            edits[row.values["ID_LINHA"]] = edit
    path = _write_and_edit(tmp_path, batch, edits)
    txt_out = tmp_path / "saida.txt"
    result = generate_cnab(
        str(path), output_path=str(txt_out), log_directory=str(tmp_path / "logs")
    )
    assert result.detail_count == 5
    assert txt_out.exists()


def test_composition_satellite_internal_credito_nao_localizado_does_not_block_generation(
    tmp_path,
):
    # Reproduz exatamente o caso relatado: o satélite nunca tem crédito
    # próprio (matching comum, isolado, nunca encontra nada para ele) mas o
    # crédito compartilhado da composição PROPOSTA já preenche ID_CREDITO/
    # LINHA_ANALITICO/REFERENCIA_ANALITICO - o rótulo interno não pode soar
    # como pendência real nem bloquear a geração quando a composição está
    # fechada (diferença 0,00) e aprovada em todas as linhas. Documento do
    # satélite deliberadamente diferente (e inválido) do documento do
    # crédito: no PFMI real (Teste 05/07/08) o beneficiário do satélite é uma
    # pessoa distinta do principal, então nem o fallback por documento o
    # encontra - usar o mesmo VALID_CPF padrão dos dois lados mascararia
    # esse caso (o satélite acabaria casando por documento "por acidente").
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00", doc="39053344705")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="1000.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert comp["diferenca"] == Decimal("0.00")
    rows = _composition_rows(batch, comp["composicao_id"])
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == REPRESENTANTE)
    principal_row = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL)
    # O satélite nunca encontra crédito próprio isoladamente, mas o rótulo
    # interno não é mais a pendência "não localizado" e o crédito
    # compartilhado já está presente.
    assert satellite.values["MOTIVO_CORRESPONDENCIA"] == "COMPOSICAO_SATELITE_CREDITO_COMPARTILHADO"
    assert satellite.values["ID_CREDITO"] == principal_row.values["ID_CREDITO"]
    assert satellite.values["ID_CREDITO"]
    assert "CREDITO_NAO_LOCALIZADO" not in satellite.values["PENDENCIAS"]
    assert "SELECAO_MANUAL_NECESSARIA" not in satellite.values["PENDENCIAS"]
    assert "DIVERGENCIA_VALOR" not in principal_row.values["PENDENCIAS"]

    edits = {
        principal_row.values["ID_LINHA"]: {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 600.00,
        },
        satellite.values["ID_LINHA"]: {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 400.00,
            "APROVADO": "SIM",
            "NOME_CEDENTE": "REPRESENTANTE FICTICIO",
            "DOC_CEDENTE": VALID_CPF,
            "TIPO_PESSOA_CEDENTE": 1,
        },
    }
    path = _write_and_edit(tmp_path, batch, edits)
    txt_out = tmp_path / "saida.txt"
    result = generate_cnab(
        str(path), output_path=str(txt_out), log_directory=str(tmp_path / "logs")
    )
    assert result.detail_count == 2
    assert txt_out.exists()


def test_composition_real_value_divergence_still_blocked_despite_shared_credit(tmp_path):
    # Confirma que a correção não libera divergência real de valor: quando a
    # soma não fecha (diferenca != 0), o estado nunca é PROPOSTA e os rótulos
    # internos continuam protegidos como antes - a composição segue bloqueada.
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "39.99")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_DIVERGENCIA_VALOR"
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = {}
    for row in rows:
        edits[row.values["ID_LINHA"]] = {"COMPOSICAO_APROVADA": "SIM", "INCLUIR_CNAB": "SIM"}
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("só é permitido quando" in issue for issue in excinfo.value.issues)


# Redução segura de pendências manuais (satélites e TIPO_PESSOA_CEDENTE) ----
# Todos os dados são fictícios. Nenhuma mudança em CONEX, comissão, matching
# por similaridade, eligible, auto-select já homologado ou regra de múltiplos
# créditos - ver decisões registradas na conversa.

SATELITE_CPF_PROPRIO = "39053344705"
SATELITE_CNPJ_PROPRIO = "11444777000161"


def _satellite_payment(cedent, value, *, beneficiary_name, beneficiary_document, block="B0001"):
    return PfmiPayment(
        block_id=block,
        source_row=1,
        title="FUNDO",
        failure=FALENCIA,
        cedent_name=cedent,
        beneficiary_name=beneficiary_name,
        beneficiary_document=beneficiary_document,
        operation_value=Decimal(value),
        beneficiary_value=Decimal(value),
    )


def test_tipo_pessoa_cedente_cpf_valido_gera_tipo_1():
    batch = _prepare(
        [_pay(PRINCIPAL, "100.00", doc=VALID_CPF)],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", doc=VALID_CPF)],
    )
    row = batch.rows[0].values
    assert row["DOC_CEDENTE"] == VALID_CPF
    assert row["TIPO_PESSOA_CEDENTE"] == 1


def test_tipo_pessoa_cedente_cnpj_valido_gera_tipo_2():
    batch = _prepare(
        [_pay(PRINCIPAL, "100.00", doc=VALID_CNPJ)],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", doc=VALID_CNPJ)],
    )
    row = batch.rows[0].values
    assert row["DOC_CEDENTE"] == VALID_CNPJ
    assert row["TIPO_PESSOA_CEDENTE"] == 2


def test_tipo_pessoa_cedente_documento_invalido_nao_gera_tipo():
    # CPF com 11 dígitos mas dígito verificador inválido: nunca decide só
    # pelo tamanho.
    invalido = "12345678901"
    batch = _prepare(
        [_pay(PRINCIPAL, "100.00", doc=invalido)],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", doc=invalido)],
    )
    row = batch.rows[0].values
    assert row["TIPO_PESSOA_CEDENTE"] == ""
    assert "DOC_CEDENTE_INVALIDO" in row["PENDENCIAS"]


def test_satellite_reuses_own_pfmi_document_when_valid():
    # O satélite tem, no próprio componente da PFMI, um documento válido e
    # um nome já registrados - nunca vindos de outro crédito do Analítico.
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE PROPRIO FICTICIO",
                beneficiary_document=SATELITE_CPF_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    rows = _composition_rows(batch, comp["composicao_id"])
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == REPRESENTANTE)
    principal = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL)
    assert satellite.values["NOME_CEDENTE"] == "REPRESENTANTE PROPRIO FICTICIO"
    assert satellite.values["DOC_CEDENTE"] == SATELITE_CPF_PROPRIO
    assert satellite.values["TIPO_PESSOA_CEDENTE"] == 1
    assert "DOC_CEDENTE_INVALIDO" not in satellite.values["PENDENCIAS"]
    # nunca associa um crédito analítico independente ao satélite
    assert satellite.values["ID_CREDITO"] == principal.values["ID_CREDITO"]
    assert satellite.values["LINHA_ANALITICO"] == principal.values["LINHA_ANALITICO"]
    assert satellite.values["REFERENCIA_ANALITICO"] == principal.values["REFERENCIA_ANALITICO"]
    assert satellite.values["NOME_CEDENTE_ANALITICO"] == principal.values["NOME_CEDENTE_ANALITICO"]
    # aprovação de nome/composição continuam manuais mesmo com o preenchimento
    assert satellite.values["APROVADO"] == ""
    assert satellite.values["COMPOSICAO_APROVADA"] == ""
    assert satellite.values["INCLUIR_CNAB"] == "NAO"
    # auditoria registra o preenchimento com a fonte exata
    auditoria = json.loads(satellite.values["PREENCHIMENTOS_AUTOMATICOS"])
    entrada = next(e for e in auditoria if e["campo"] == "NOME_CEDENTE/DOC_CEDENTE")
    assert "PFMI" in entrada["fonte"]
    assert entrada["regra"] == "PROPRIO_COMPONENTE_PFMI_DOCUMENTO_CHECKSUM_VALIDO"


def test_satellite_does_not_reuse_document_when_own_document_invalid():
    # Documento inválido no próprio componente: mantém DOC_CEDENTE_INVALIDO,
    # nunca busca nem inventa outro documento por nome.
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE FICTICIO SEM DOC",
                beneficiary_document="00000000000",
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == REPRESENTANTE)
    assert satellite.values["NOME_CEDENTE"] == ""
    assert satellite.values["DOC_CEDENTE"] == ""
    assert satellite.values["TIPO_PESSOA_CEDENTE"] == ""
    assert "DOC_CEDENTE_INVALIDO" in satellite.values["PENDENCIAS"]
    assert satellite.values["PREENCHIMENTOS_AUTOMATICOS"] == ""


def test_satellite_never_uses_independent_analytic_credit_even_if_matched_alone():
    # Mesmo quando o próprio pagamento do satélite coincide, por nome E
    # documento, com um crédito INDEPENDENTE do Analítico (não o crédito
    # compartilhado da composição), o satélite nunca troca de crédito, nunca
    # duplica, e o nominal/valor de aquisição da composição não mudam.
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name=REPRESENTANTE,
                beneficiary_document=VALID_CPF,
            ),
        ],
        [
            _credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00"),
            # crédito independente que bateria por nome+documento com o
            # satélite se o sistema procurasse (não deve ser usado)
            _credit(
                "REF-2", REPRESENTANTE, acquisition="999.00", nominal="999.00", doc=VALID_CPF
            ),
        ],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "PROPOSTA"
    assert comp["total_analitico"] == Decimal("100.00")
    assert comp["nominal_analitico"] == Decimal("500.00")
    rows = _composition_rows(batch, comp["composicao_id"])
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == REPRESENTANTE)
    principal = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == PRINCIPAL)
    assert satellite.values["ID_CREDITO"] == principal.values["ID_CREDITO"]
    assert satellite.values["REFERENCIA_ANALITICO"] == "REF-1"
    assert satellite.values["NOME_CEDENTE_ANALITICO"] == PRINCIPAL
    assert satellite.values["AQUISICAO_ORIGINAL"] == Decimal("100.00")
    assert satellite.values["NOMINAL_ORIGINAL"] == Decimal("500.00")
    # a soma de créditos selecionados nunca duplica (dedup por identidade)
    assert len({r.values["ID_CREDITO"] for r in rows}) == 1


def test_no_suggestion_confirms_identity_automatically():
    # Nenhuma das sugestões novas seleciona ou aprova nada sozinha. Usa o
    # padrão 60/40 (como as demais composições fictícias deste arquivo): o
    # principal sozinho (60.00) não fecha com a aquisição do crédito
    # (100.00), então nem o auto-select por nome já homologado (que depende
    # só do total do PRÓPRIO grupo, nunca da composição) se aplicaria aqui -
    # ver decisão R1, preservada e não afetada por esta mudança.
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00", doc=VALID_CPF),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="SATELITE VALIDO FICTICIO",
                beneficiary_document=SATELITE_CNPJ_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", doc=VALID_CPF)],
    )
    assert batch.compositions[0]["estado"] == "PROPOSTA"
    for row in batch.rows:
        if row.values["COMPOSICAO_ID"]:
            assert row.values["INCLUIR_CNAB"] == "NAO"
            assert row.values["COMPOSICAO_APROVADA"] == ""


def test_composition_still_requires_composicao_aprovada_with_zero_difference():
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE PROPRIO FICTICIO",
                beneficiary_document=SATELITE_CPF_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00")],
    )
    comp = batch.compositions[0]
    assert comp["diferenca"] == Decimal("0.00")
    rows = _composition_rows(batch, comp["composicao_id"])
    assert all(r.values["COMPOSICAO_APROVADA"] == "" for r in rows)


def test_composition_nominal_sum_one_cent_difference_blocks_generation(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = {}
    for row in rows:
        is_principal = row.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edit = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            # soma = 500.01, um centavo a mais que o nominal do crédito (500.00)
            "VL_NOMINAL": 300.01 if is_principal else 200.00,
        }
        if not is_principal:
            edit.update(
                {"APROVADO": "SIM", "NOME_CEDENTE": "REPRESENTANTE FICTICIO",
                 "DOC_CEDENTE": VALID_CPF, "TIPO_PESSOA_CEDENTE": 1}
            )
        edits[row.values["ID_LINHA"]] = edit
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    issues = excinfo.value.issues
    assert any(
        "soma dos nominais informados" in issue and "diferença de +0.01" in issue
        for issue in issues
    )


def test_composition_nominal_sum_exact_generates_successfully(tmp_path):
    batch = _prepare(
        [_pay(PRINCIPAL, "60.00"), _pay(REPRESENTANTE, "40.00")],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = {}
    for row in rows:
        is_principal = row.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edit = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 300.00 if is_principal else 200.00,
        }
        if not is_principal:
            edit.update(
                {"APROVADO": "SIM", "NOME_CEDENTE": "REPRESENTANTE FICTICIO",
                 "DOC_CEDENTE": VALID_CPF, "TIPO_PESSOA_CEDENTE": 1}
            )
        edits[row.values["ID_LINHA"]] = edit
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    validated = validate_for_generation(loaded)
    assert len(validated.rows) == 2


def test_composition_nominal_suggestion_preserves_human_approval():
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE PROPRIO FICTICIO",
                beneficiary_document=SATELITE_CPF_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="777.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    assert [r.values["VL_NOMINAL"] for r in rows] == [Decimal("466.20"), Decimal("310.80")]
    assert all(r.values["COMPOSICAO_APROVADA"] == "" for r in rows)
    assert all("PREENCHER_VL_NOMINAL_COMPOSICAO" not in r.values["PENDENCIAS"] for r in rows)


def test_satellite_divergencia_nome_still_requires_aprovado(tmp_path):
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE PROPRIO FICTICIO",
                beneficiary_document=SATELITE_CPF_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00")],
    )
    comp = batch.compositions[0]
    rows = _composition_rows(batch, comp["composicao_id"])
    satellite = next(r for r in rows if r.values["NOME_CEDENTE_PFMI"] == REPRESENTANTE)
    assert "DIVERGENCIA_NOME" in satellite.values["PENDENCIAS"]
    edits = {}
    for row in rows:
        is_principal = row.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edits[row.values["ID_LINHA"]] = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 300.00 if is_principal else 200.00,
            # propositalmente sem APROVADO=SIM no satélite
        }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("DIVERGENCIA_NOME" in issue for issue in excinfo.value.issues)


def test_manifest_and_protection_remain_valid_with_new_audit_field(tmp_path):
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE PROPRIO FICTICIO",
                beneficiary_document=SATELITE_CPF_PROPRIO,
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00")],
    )
    from openpyxl import load_workbook

    path = tmp_path / "intermediario.xlsx"
    write_intermediate(batch, path)
    loaded = read_intermediate(path)
    assert loaded.structural_issues == []
    book = load_workbook(path)
    sheet = book["CREDITOS"]
    header = {cell.value: cell.column for cell in sheet[1]}
    sheet.cell(2, header["PREENCHIMENTOS_AUTOMATICOS"]).value = "[]"
    book.save(path)
    book.close()
    tampered = read_intermediate(path)
    assert any("PREENCHIMENTOS_AUTOMATICOS" in issue for issue in tampered.structural_issues)


def test_generation_still_refused_when_satellite_document_unresolved(tmp_path):
    # Documento inválido em qualquer componente já bloqueia a composição
    # inteira na detecção (BLOQUEADA_DOCUMENTO_INVALIDO, regra já existente,
    # não relacionada a esta mudança) - a geração continua recusada.
    batch = _prepare(
        [
            _pay(PRINCIPAL, "60.00"),
            _satellite_payment(
                REPRESENTANTE, "40.00",
                beneficiary_name="REPRESENTANTE SEM DOC",
                beneficiary_document="00000000000",
            ),
        ],
        [_credit("REF-1", PRINCIPAL, acquisition="100.00", nominal="500.00")],
    )
    comp = batch.compositions[0]
    assert comp["estado"] == "BLOQUEADA_DOCUMENTO_INVALIDO"
    rows = _composition_rows(batch, comp["composicao_id"])
    edits = {}
    for row in rows:
        is_principal = row.values["NOME_CEDENTE_PFMI"] == PRINCIPAL
        edits[row.values["ID_LINHA"]] = {
            "COMPOSICAO_APROVADA": "SIM",
            "INCLUIR_CNAB": "SIM",
            "VL_NOMINAL": 300.00 if is_principal else 200.00,
        }
    path = _write_and_edit(tmp_path, batch, edits)
    loaded = read_intermediate(path)
    with pytest.raises(ValidationError) as excinfo:
        validate_for_generation(loaded)
    assert any("só é permitido quando" in issue for issue in excinfo.value.issues)
