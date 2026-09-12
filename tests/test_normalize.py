from __future__ import annotations

from decimal import Decimal

import pytest

from gerador_cnab_nox.normalize import (
    digits,
    money,
    normalize_name,
    parse_date,
    parse_positive_int,
    validate_document,
)


def test_explicit_name_equivalences() -> None:
    assert normalize_name("Companhia Árvore Sociedade Anônima") == normalize_name("CIA ARVORE S/A")
    assert normalize_name("Acme Limitada") == normalize_name("ACME LTDA.")
    assert normalize_name("Acme Sociedade Anônima") == normalize_name("ACME S.A.")
    assert normalize_name("Acme Companhia") == normalize_name("ACME C.I.A.")
    assert normalize_name("Acme Limitada") == normalize_name("ACME L.T.D.A.")
    assert normalize_name("MARIA E JOAO") != normalize_name("MARIA JOAO")


def test_documents_and_integral_float() -> None:
    assert digits(52998224725.0) == "52998224725"
    assert validate_document("529.982.247-25") == ("52998224725", 1)
    assert validate_document("11.222.333/0001-81") == ("11222333000181", 2)
    assert validate_document("111.111.111-11")[1] is None


def test_money_exact_cents() -> None:
    assert money("R$ 1.234,565") == Decimal("1234.57")
    assert money(0.01) == Decimal("0.01")
    with pytest.raises(ValueError):
        money("não é valor")
    with pytest.raises(ValueError):
        money("1e999999")


def test_extreme_date_and_sequence_are_validation_errors() -> None:
    with pytest.raises(ValueError):
        parse_date(999_999_999)
    with pytest.raises(ValueError):
        parse_positive_int(float("inf"))
