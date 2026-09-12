from datetime import date
from decimal import Decimal

from gerador_cnab_nox.matching import _group_payments
from gerador_cnab_nox.models import AnalyticCredit, PfmiPayment
from scripts.diagnose_reconciliation import name_variants, suggest


def group(name="MARIA SILVA", amount="100", issues=()):
    payment = PfmiPayment(
        "B1", 7, "", "FALENCIA A", name, "TERCEIRO", "", Decimal(amount),
        Decimal(amount), issues=issues,
    )
    return _group_payments([payment])[0]


def credit(row=2, name="MARIA SILVA", amount="100", campaign="FALENCIA A"):
    return AnalyticCredit(
        row, str(row), "", name, campaign, Decimal("200"), Decimal(amount), date(2026, 1, 1),
    )


def test_value_alone_does_not_suggest_an_unrelated_person():
    assert suggest(group(), [credit(name="JOAO PEREIRA")]) == []


def test_one_cent_difference_is_not_a_suggestion():
    assert suggest(group(), [credit(amount="100.01")]) == []


def test_ambiguous_candidates_are_preserved_and_not_approved():
    results = suggest(group(), [credit(), credit(row=3)])
    assert [r["linha_analitico"] for r in results] == [2, 3]
    assert all(r["acao"] == "CONFIRMAR_IDENTIDADE" for r in results)


def test_unknown_campaign_is_explicit_not_a_new_alias():
    result = suggest(group(), [credit(campaign="OUTRA CAMPANHA")])[0]
    assert result["falencia_equivalente_confirmada"] is False
    assert result["acao"] == "CONFIRMAR_IDENTIDADE_E_FALENCIA"


def test_estate_and_parentheses_are_search_hints_only():
    assert "MARIA SILVA" in name_variants("Espólio de Maria Silva")
    assert "MARIA SILVA" in name_variants("Advogado (Maria Silva)")
    assert suggest(group("Espólio de Maria Silva"), [credit()])[0]["acao"] == "CONFIRMAR_IDENTIDADE"


def test_invalid_source_has_no_suggestions():
    assert suggest(group(issues=("VALOR_INVALIDO",)), [credit()]) == []


def test_does_not_combine_credits_to_match_total():
    assert suggest(group(), [credit(amount="40"), credit(row=3, amount="60")]) == []
