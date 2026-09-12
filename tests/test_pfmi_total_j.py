from dataclasses import replace
from decimal import Decimal

from openpyxl import Workbook

from gerador_cnab_nox.matching import _group_payments
from gerador_cnab_nox.readers import read_pfmi


def test_j_drives_group_and_summary_even_when_b_differs(tmp_path):
    path = tmp_path / "pfmi.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["DATA", "QUANTIDADE DE TEDS", "VALOR TOTAL A PAGAR"])
    sheet.append(["10/09/2026", 3, 11701.10])
    sheet.append([])
    sheet.append([
        "", "VALOR DA OPERAÇÃO", "FALÊNCIA", "CREDOR / CEDENTE",
        "BENEFICIÁRIOS / FAVORECIDOS", "BCO", "AG", "CTA", "CPF/CNPJ", "VALOR",
    ])
    for amount in (9360.88, 1170.11, 1170.11):
        sheet.append(
            ["", 999, "FALENCIA A", "CREDOR TESTE", "BENEFICIARIO", "", "", "", "", amount]
        )
    workbook.save(path)
    workbook.close()
    data = read_pfmi(path)
    assert _group_payments(data.payments)[0].total == Decimal("11701.10")
    assert any(e.get("alerta") == "RESUMO_CONFERIDO" for e in data.evidence)
    assert all("DIVERGENCIA_VALOR_LINHA" in p.issues for p in data.payments)
    # A missing J must not silently fall back to B, and zero J is a real zero.
    first = data.payments[0]
    assert _group_payments([replace(first, beneficiary_value=None)])[0].total is None
    assert _group_payments([replace(first, beneficiary_value=Decimal(0))])[0].total == 0
