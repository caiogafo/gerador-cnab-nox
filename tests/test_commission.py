from dataclasses import replace
from datetime import date
from decimal import Decimal

import pytest

from gerador_cnab_nox.matching import prepare_batch
from gerador_cnab_nox.models import CESSAO, PfmiData
from gerador_cnab_nox.normalize import calculate_commission

from .test_matching import _credit, _due, _payment


@pytest.mark.parametrize(
    ("n", "a", "c", "e"),
    [
        ("181800.00", "69993.00", "16771.05", "86764.05"),
        ("25594.84", "10807.55", "2218.09", "13025.64"),
        ("23191.21", "11767.38", "1713.57", "13480.95"),
        ("181800.00", "95000.00", "13020.00", "108020.00"),
        ("16588.17", "4459.95", "1819.23", "6279.18"),
    ],
)
def test_t04_reference_per_credit(n, a, c, e):
    _, commission, expected = calculate_commission(n, a, Decimal("15"))
    assert (commission, expected) == (Decimal(c), Decimal(e))


def test_t04_half_cent_per_credit_not_aggregate():
    first = calculate_commission("1.05", "1", Decimal("10"))
    second = calculate_commission("1.05", "1", Decimal("10"))
    assert first[1] == second[1] == Decimal("0.01")
    assert first[1] + second[1] == Decimal("0.02")
    assert calculate_commission("2.10", "2", Decimal("10"))[1] == Decimal("0.01")
    with pytest.raises(ValueError, match="COMISSAO_BASE_NEGATIVA"):
        calculate_commission("90", "100", Decimal("15"))
    assert calculate_commission("90", "100", Decimal("0"))[2] == Decimal("100")


def test_t02_t13_t15_independent_modalities():
    payment = _payment("B1", "107.50")
    payment = replace(
        payment,
        modality=CESSAO,
        commission_text="15",
        beneficiary_name="CONEXCRED",
        beneficiary_document="11222333000181",
    )
    batch = prepare_batch(
        PfmiData(payments=[payment]),
        [_credit("A", "100")],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    row = batch.rows[0]
    assert row.values["INCLUIR_CNAB"] == "SIM"
    assert row.values["NOME_CEDENTE"] == "CONEXCRED INTERMEDIACAO"
    assert row.values["DOC_CEDENTE"] == "11222333000181"
    assert row.values["DT_EMISSAO_TITULO"] is None
    assert row.values["ASSINATURA_ORIGINAL"] == date(2026, 9, 1)
    assert row.values["AQUISICAO_ORIGINAL"] == Decimal("100")
    assert row.values["COMISSAO_CALCULADA"] == Decimal("7.50")
    assert row.values["VL_PRESENTE"] == Decimal("107.50")
    assert row.values["TAXA_INDEXADOR"] == ""


def test_t04_reference_total_is_sum_of_rounded_credits():
    pairs = [
        ("181800.00", "69993.00"),
        ("25594.84", "10807.55"),
        ("23191.21", "11767.38"),
        ("181800.00", "95000.00"),
        ("16588.17", "4459.95"),
    ]
    results = [calculate_commission(n, a, Decimal("15")) for n, a in pairs]
    assert sum(r[1] for r in results) == Decimal("35541.94")
    assert sum(r[2] for r in results) == Decimal("227569.82")
    assert calculate_commission("428974.22", "192027.88", Decimal("15"))[1] == Decimal("35541.95")
