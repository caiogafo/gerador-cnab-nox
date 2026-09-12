from __future__ import annotations

from copy import deepcopy
from datetime import date
from decimal import Decimal

import pytest

from gerador_cnab_nox.errors import ValidationError
from gerador_cnab_nox.models import PreparedRow
from gerador_cnab_nox.validation import validate_for_generation
from gerador_cnab_nox.workbook import LoadedIntermediate

from .conftest import VALID_CNPJ, VALID_CPF


def _prepared(
    line_id: str,
    group: str,
    *,
    analytic_line: str = "2",
    reference: str = "P-1",
) -> PreparedRow:
    values = {
        "ID_LINHA": line_id,
        "ID_BLOCO": "B1",
        "ID_GRUPO": group,
        "ID_CREDITO": f"snapshot|{analytic_line}|{reference}" if analytic_line else "",
        "LINHA_ANALITICO": analytic_line,
        "NOME_CEDENTE_PFMI": "ACME LTDA",
        "REFERENCIA_ANALITICO": reference,
        "TOTAL_PFMI_GRUPO": Decimal("100.00"),
        "INCLUIR_CNAB": "SIM",
        "APROVADO": "",
        "DOC_CEDENTE": VALID_CPF,
        "NOME_CEDENTE": "ACME LTDA",
        "SEU_NUMERO": "1",
        "NU_DOCUMENTO": "1",
        "DT_VENCIMENTO": date(2030, 1, 31),
        "VL_NOMINAL": Decimal("150.00"),
        "DOC_SACADO": VALID_CNPJ,
        "NOME_SACADO": "FALENCIA ALFA SA",
        "VL_PRESENTE": Decimal("100.00"),
        "TIPO_PESSOA_SACADO": 2,
        "ENDERECO": "",
        "CEP": "",
        "TP_TITULO": 24,
        "DT_EMISSAO_TITULO": date(2026, 9, 1),
        "COOBRIGACAO": 2,
        "TIPO_PESSOA_CEDENTE": 1,
        "NFE": "",
        "VALOR_PAGO_TITULO": Decimal("0.00"),
        "INDEXADOR": "",
        "TAXA_INDEXADOR": "",
        "MOVIMENTO": 1,
    }
    originals = deepcopy(values)
    originals.update(
        {
            "TOTAL_PFMI_GRUPO": Decimal("100.00"),
            "NOME_CEDENTE_PFMI": "ACME LTDA",
            "LINHA_ANALITICO": analytic_line,
            "REFERENCIA_ANALITICO": reference,
        }
    )
    return PreparedRow(values, originals)


def _loaded(rows: list[PreparedRow]) -> LoadedIntermediate:
    return LoadedIntermediate(
        rows=rows,
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
        expected_ids=[row.values["ID_LINHA"] for row in rows],
        unknown_count=0,
        structural_issues=[],
    )


def test_same_analytic_credit_cannot_be_selected_twice() -> None:
    rows = [_prepared("A", "G1"), _prepared("B", "G2")]
    with pytest.raises(ValidationError, match="mesmo crédito"):
        validate_for_generation(_loaded(rows))


def test_placeholder_without_analytic_origin_cannot_be_selected() -> None:
    row = _prepared("A", "G1", analytic_line="", reference="")
    with pytest.raises(ValidationError, match="sem crédito original"):
        validate_for_generation(_loaded([row]))


def test_formula_and_non_integral_type_are_rejected() -> None:
    formula = _prepared("A", "G1")
    formula.values["VL_PRESENTE"] = "=50+50"
    with pytest.raises(ValidationError, match="fórmulas"):
        validate_for_generation(_loaded([formula]))

    non_integral = _prepared("A", "G1")
    non_integral.values["TIPO_PESSOA_CEDENTE"] = 1.5
    with pytest.raises(ValidationError, match="TIPO_PESSOA_CEDENTE"):
        validate_for_generation(_loaded([non_integral]))


def test_sequence_fields_are_recalculated_and_reported() -> None:
    row = _prepared("A", "G1")
    row.values["SEU_NUMERO"] = 999
    row.values["NU_DOCUMENTO"] = 888
    validated = validate_for_generation(_loaded([row]))
    assert validated.rows[0]["SEU_NUMERO"] == 1
    assert validated.rows[0]["NU_DOCUMENTO"] == 1
    assert len(validated.warnings) >= 2


@pytest.mark.parametrize("field", ["NOME_CEDENTE", "NOME_SACADO"])
def test_name_without_alphanumeric_characters_is_rejected(field: str) -> None:
    row = _prepared("A", "G1")
    row.values[field] = "---"
    row.values["APROVADO"] = "SIM"

    with pytest.raises(ValidationError, match=field):
        validate_for_generation(_loaded([row]))
