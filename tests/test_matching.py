from __future__ import annotations

from datetime import date
from decimal import Decimal

from gerador_cnab_nox.matching import prepare_batch
from gerador_cnab_nox.models import AnalyticCredit, DueRecord, PfmiData, PfmiPayment

from .conftest import VALID_CNPJ, VALID_CPF


def _payment(block: str, value: str) -> PfmiPayment:
    return PfmiPayment(
        block_id=block,
        source_row=1,
        title="FUNDO",
        failure="FALENCIA ALFA SA",
        cedent_name="ACME LTDA",
        beneficiary_name="BENEFICIARIO",
        beneficiary_document=VALID_CPF,
        operation_value=Decimal(value),
        beneficiary_value=Decimal(value),
    )


def _credit(reference: str, value: str) -> AnalyticCredit:
    return AnalyticCredit(
        source_row=2,
        reference=reference,
        cedent_document=VALID_CPF,
        cedent_name="ACME LIMITADA",
        campaign="FALENCIA ALFA SOCIEDADE ANONIMA",
        nominal_value=Decimal("150.00"),
        acquisition_value=Decimal(value),
        signature_date=date(2026, 9, 1),
    )


def _due() -> DueRecord:
    return DueRecord(2, "FALENCIA ALFA S/A", VALID_CNPJ, date(2030, 1, 1))


def test_exact_candidate_set_requires_manual_selection_when_multiple_principals() -> None:
    # Security fix: two or more principal credits (even same document, exact sum)
    # are never auto-selected; both go to manual review with a plain message.
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "40.00"), _payment("B0001", "60.00")]),
        [_credit("A", "40.00"), _credit("B", "60.00")],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=10,
    )
    assert [row.values["INCLUIR_CNAB"] for row in batch.rows] == ["NAO", "NAO"]
    assert all("MULTIPLOS_CREDITOS_PRINCIPAIS" in row.values["PENDENCIAS"] for row in batch.rows)
    assert all(
        row.values["ACAO_NECESSARIA"]
        == "Encontramos mais de um crédito. Selecione os créditos corretos."
        for row in batch.rows
    )


def test_invalid_cedent_document_is_never_auto_selected() -> None:
    invalid = AnalyticCredit(**{**_credit("A", "100.00").__dict__, "cedent_document": "123"})
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [invalid],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=10,
    )

    assert batch.rows[0].values["INCLUIR_CNAB"] == "NAO"
    assert "DOC_CEDENTE_INVALIDO" in batch.rows[0].values["PENDENCIAS"]


def test_missing_acquisition_value_is_never_auto_selected() -> None:
    missing_value = AnalyticCredit(**{**_credit("A", "0.00").__dict__, "acquisition_value": None})
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "0.00")]),
        [missing_value],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=10,
    )

    assert batch.rows[0].values["INCLUIR_CNAB"] == "NAO"
    assert "VL_PRESENTE_INVALIDO" in batch.rows[0].values["PENDENCIAS"]


def test_negative_candidate_values_are_visible_and_not_auto_selected() -> None:
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [_credit("A", "-10.00"), _credit("B", "110.00")],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=10,
    )

    assert [row.values["INCLUIR_CNAB"] for row in batch.rows] == ["NAO", "NAO"]
    assert all(row.values["STATUS"] == "REVISAR" for row in batch.rows)
    assert "VL_PRESENTE_NEGATIVO" in batch.rows[0].values["PENDENCIAS"]


def test_one_cent_difference_requires_review(monkeypatch) -> None:
    monkeypatch.setenv("MAX_TOLERANCE_DIFF_REAIS", "0")
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [_credit("A", "99.99")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert batch.rows[0].values["INCLUIR_CNAB"] == "NAO"
    assert batch.rows[0].values["STATUS"] == "DIVERGENCIA_VALOR"


def test_same_person_in_two_physical_blocks_is_not_merged() -> None:
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00"), _payment("B0002", "100.00")]),
        [_credit("A", "100.00")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len({row.values["ID_GRUPO"] for row in batch.rows}) == 2
    assert [row.values["INCLUIR_CNAB"] for row in batch.rows] == ["NAO", "NAO"]
    assert all(
        "CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS" in row.values["PENDENCIAS"] for row in batch.rows
    )


def test_conflicting_due_records_are_visible() -> None:
    conflict = DueRecord(3, "FALENCIA ALFA SA", VALID_CPF, date(2030, 1, 2))
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [_credit("A", "100.00")],
        [_due(), conflict],
        liquidation_date=None,
        first_sequence=None,
    )
    assert batch.rows[0].values["STATUS"] == "CONFLITO_BASE_VENCIMENTOS"
    assert batch.rows[0].values["DOC_SACADO"] == ""


def test_invalid_due_data_is_visible_during_preparation() -> None:
    invalid = DueRecord(2, "FALENCIA ALFA SA", "123", None)
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [_credit("A", "100.00")],
        [invalid],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    row = batch.rows[0].values
    assert row["STATUS"] == "PREENCHER_MANUALMENTE"
    assert "DOC_SACADO_INVALIDO" in row["PENDENCIAS"]
    assert "DATA_VENCIMENTO_INVALIDA" in row["PENDENCIAS"]


def test_signature_after_liquidation_is_visible_during_preparation() -> None:
    credit = _credit("A", "100.00")
    future_credit = AnalyticCredit(**{**credit.__dict__, "signature_date": date(2026, 9, 4)})
    batch = prepare_batch(
        PfmiData(payments=[_payment("B0001", "100.00")]),
        [future_credit],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    assert "DT_EMISSAO_POSTERIOR_LIQUIDACAO" in batch.rows[0].values["PENDENCIAS"]


def test_t08_four_payments_one_credit_and_no_subset_search():
    payments = [_payment("B1", "25") for _ in range(4)]
    b = prepare_batch(
        PfmiData(payments=payments),
        [_credit("A", "100")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len(b.rows) == 1 and b.rows[0].values["INCLUIR_CNAB"] == "SIM"
    assert b.payment_count == 4 and b.group_count == 1
    b = prepare_batch(
        PfmiData(payments=payments),
        [_credit("A", "40"), _credit("B", "60"), _credit("C", "200")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len(b.rows) == 3 and all(r.values["INCLUIR_CNAB"] == "NAO" for r in b.rows)


def test_t09_repeated_references_and_homonyms_are_distinct():
    from dataclasses import replace

    first = _credit("REPEAT", "50")
    second = replace(first, source_row=3)
    b = prepare_batch(
        PfmiData(payments=[_payment("B1", "100")]),
        [first, second],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len({r.values["ID_CREDITO"] for r in b.rows}) == 2
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in b.rows)
    assert all("MULTIPLOS_CREDITOS_PRINCIPAIS" in r.values["PENDENCIAS"] for r in b.rows)
    second = replace(second, cedent_document=VALID_CNPJ)
    b = prepare_batch(
        PfmiData(payments=[_payment("B1", "100")]),
        [first, second],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in b.rows)
    assert all("DOCUMENTOS_CANDIDATOS_CONFLITANTES" in r.values["PENDENCIAS"] for r in b.rows)


def test_t11_mogiano_three_sources_and_no_fuzzy_failure():
    from dataclasses import replace

    payment = replace(_payment("B1", "100"), failure="MOGIANO")
    credit = replace(_credit("A", "100"), campaign="MOGIANO TRANSP GERAIS")
    due = replace(_due(), debtor_name="MOGIANO")
    b = prepare_batch(
        PfmiData(payments=[payment]), [credit], [due], liquidation_date=None, first_sequence=None
    )
    assert (
        b.rows[0].values["INCLUIR_CNAB"] == "SIM" and b.rows[0].values["DOC_SACADO"] == VALID_CNPJ
    )
    bad = replace(credit, campaign="MOGIANO TRANSPORTES GERAIS")
    b = prepare_batch(
        PfmiData(payments=[payment]), [bad], [due], liquidation_date=None, first_sequence=None
    )
    assert b.rows[0].values["ID_CREDITO"] == ""


def test_keleti_quirino_matches_canonical_analytic_failure():
    from dataclasses import replace

    name = "QUIRINO DANTAS DA ROCHA"
    payment = replace(_payment("B1", "100"), failure="  keleti  ", cedent_name=name)
    credit = replace(_credit("A", "100"), campaign="KELETI ENGENHARIA", cedent_name=name)
    due = replace(_due(), debtor_name="KELETI ENGENHARIA")
    batch = prepare_batch(
        PfmiData(payments=[payment]), [credit], [due],
        liquidation_date=date(2026, 9, 17), first_sequence=1,
    )
    row = batch.rows[0].values
    assert row["STATUS"] == "OK"
    assert row["INCLUIR_CNAB"] == "SIM"
    assert row["ID_CREDITO"]
    assert not row["PENDENCIAS"]


def test_t12_individual_alternative_requires_manual_selection():
    from dataclasses import replace

    other = replace(_credit("B", "100"), cedent_name="OUTRA PESSOA")
    similar = replace(_credit("C", "99.99"), cedent_name="ACME LTDA X")
    b = prepare_batch(
        PfmiData(payments=[_payment("B1", "100")]),
        [other, similar],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len(b.rows) == 1
    r = b.rows[0].values
    assert r["MOTIVO_CORRESPONDENCIA"] == "ALTERNATIVA_VALOR_EXATO_ESCOLHA_MANUAL"
    assert r["NOME_CEDENTE"] == "OUTRA PESSOA" and r["INCLUIR_CNAB"] == "NAO"


def test_t13_normal_never_uses_beneficiary_identity_and_conex_conflict():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    normal = replace(_payment("B1", "100"), beneficiary_document=VALID_CNPJ)
    invalid = replace(_credit("A", "100"), cedent_document="123")
    b = prepare_batch(
        PfmiData(payments=[normal]), [invalid], [_due()], liquidation_date=None, first_sequence=None
    )
    assert b.rows[0].values["DOC_CEDENTE"] == "123"
    p = replace(
        normal,
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        operation_value=Decimal("50"),
        beneficiary_value=Decimal("50"),
    )
    for doc in ("", VALID_CPF):
        b = prepare_batch(
            PfmiData(payments=[p, replace(p, beneficiary_document=doc)]),
            [_credit("A", "100")],
            [_due()],
            liquidation_date=None,
            first_sequence=None,
        )
        assert b.rows[0].values["DOC_CEDENTE"] == ""
        assert b.rows[0].values["NOME_CEDENTE"] == "CONEXCRED INTERMEDIACAO"


def test_conexcred_beneficiary_recognized_by_prefix_not_exact_name():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    base = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
    )
    recognized_names = (
        "CONEXCRED",
        "CONEXCRED INTERMEDIACAO",
        "CONEXCRED INTERMEDIAÇÃO E AGENCIAMENTO DE SERVIÇOS",
    )
    for name in recognized_names:
        p = replace(base, beneficiary_name=name)
        b = prepare_batch(
            PfmiData(payments=[p]),
            [_credit("A", "100")],
            [_due()],
            liquidation_date=None,
            first_sequence=None,
        )
        row = b.rows[0].values
        assert row["ESTADO_CNPJ_CONEXCRED"] == "UNICO_VALIDO", name
        assert row["DOC_CEDENTE"] == VALID_CNPJ, name
        assert "CNPJ_CONEXCRED_AUSENTE" not in row["PENDENCIAS"], name

    # A non-Conexcred beneficiary must not be picked up by the prefix match.
    p = replace(base, beneficiary_name="OUTRA EMPRESA LTDA")
    b = prepare_batch(
        PfmiData(payments=[p]),
        [_credit("A", "100")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    row = b.rows[0].values
    assert row["ESTADO_CNPJ_CONEXCRED"] == "AUSENTE"
    assert row["DOC_CEDENTE"] == ""
    assert "CNPJ_CONEXCRED_AUSENTE" in row["PENDENCIAS"]
    assert "EMISSAO_NOVA_CESSAO_AUSENTE" in row["PENDENCIAS"]


def test_sequence_is_reserved_per_group_and_shared_by_alternatives():
    from dataclasses import replace

    pa = replace(_payment("BA", "100"), failure="FALA", cedent_name="PESSOA A")
    ca = replace(_credit("A", "100"), campaign="FALA", cedent_name="PESSOA A")
    pb = replace(_payment("BB", "100"), failure="FALB", cedent_name="PESSOA B")
    cb = replace(_credit("B", "100"), campaign="FALB", cedent_name="PESSOA B DIFERENTE")
    pc = replace(_payment("BC", "100"), failure="FALC", cedent_name="PESSOA C")
    cc = replace(_credit("C", "100"), campaign="FALC", cedent_name="PESSOA C")
    due = [
        replace(_due(), debtor_name="FALA"),
        replace(_due(), debtor_name="FALB"),
        replace(_due(), debtor_name="FALC"),
    ]
    b = prepare_batch(
        PfmiData(payments=[pa, pb, pc]),
        [ca, cb, cc],
        due,
        liquidation_date=date(2026, 9, 3),
        first_sequence=100,
    )
    by_group: dict[str, list] = {}
    for row in b.rows:
        by_group.setdefault(row.values["ID_GRUPO"], []).append(row.values)
    groups = list(by_group.values())
    assert len(groups) == 3
    # Group A: single principal, auto-selected -> sequence 100.
    assert groups[0][0]["SEU_NUMERO"] == 100
    # Group B: only a divergent-name alternative -> reserves one number too (101),
    # shared if there were more than one alternative candidate.
    assert all(r["SEU_NUMERO"] == 101 for r in groups[1])
    assert groups[1][0]["INCLUIR_CNAB"] == "SIM"
    assert groups[1][0]["APROVADO"] == "SIM"
    # Group C never shifts because of group B's alternative(s).
    assert groups[2][0]["SEU_NUMERO"] == 102


def test_cessao_emission_date_fills_every_title_of_its_own_pfmi_only():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    p1 = replace(
        _payment("B1", "40"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("40"),
        beneficiary_value=Decimal("40"),
        emissao_nova_cessao_texto="01/09/2026",
        file_id="F0001",
    )
    p2 = replace(
        _payment("B2", "60"),
        cedent_name="OUTRA PESSOA",
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("60"),
        beneficiary_value=Decimal("60"),
        emissao_nova_cessao_texto="01/09/2026",
        file_id="F0001",
    )
    p3 = replace(
        _payment("B3", "70"),
        cedent_name="TERCEIRA PESSOA",
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("70"),
        beneficiary_value=Decimal("70"),
        emissao_nova_cessao_texto="02/09/2026",
        file_id="F0002",
    )
    credits = [
        replace(_credit("A", "40"), cedent_name="ACME LTDA"),
        replace(_credit("B", "60"), cedent_name="OUTRA PESSOA"),
        replace(_credit("C", "70"), cedent_name="TERCEIRA PESSOA"),
    ]
    b = prepare_batch(
        PfmiData(payments=[p1, p2, p3]),
        credits,
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    by_name = {r.values["NOME_CEDENTE_PFMI"]: r.values for r in b.rows}
    assert by_name["ACME LTDA"]["DT_EMISSAO_TITULO"] == date(2026, 9, 1)
    assert by_name["OUTRA PESSOA"]["DT_EMISSAO_TITULO"] == date(2026, 9, 1)
    assert by_name["TERCEIRA PESSOA"]["DT_EMISSAO_TITULO"] == date(2026, 9, 2)
    for row in by_name.values():
        assert "EMISSAO_NOVA_CESSAO_AUSENTE" not in row["PENDENCIAS"]


def test_cessao_emission_date_missing_invalid_and_late_are_clear_pendencies():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    base = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
    )
    credit = _credit("A", "100")

    empty = replace(base, emissao_nova_cessao_texto="")
    b = prepare_batch(
        PfmiData(payments=[empty]), [credit], [_due()],
        liquidation_date=date(2026, 9, 3), first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] is None
    assert "EMISSAO_NOVA_CESSAO_AUSENTE" in row["PENDENCIAS"]

    invalid = replace(base, emissao_nova_cessao_texto="não é uma data")
    b = prepare_batch(
        PfmiData(payments=[invalid]), [credit], [_due()],
        liquidation_date=date(2026, 9, 3), first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] is None
    assert "EMISSAO_NOVA_CESSAO_INVALIDA" in row["PENDENCIAS"]

    late = replace(base, emissao_nova_cessao_texto="10/09/2026")
    b = prepare_batch(
        PfmiData(payments=[late]), [credit], [_due()],
        liquidation_date=date(2026, 9, 3), first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] == date(2026, 9, 10)
    assert "EMISSAO_NOVA_CESSAO_POSTERIOR_LIQUIDACAO" in row["PENDENCIAS"]


def test_normal_modality_ignores_emissao_nova_cessao_and_uses_signature():
    from dataclasses import replace

    p = replace(_payment("B1", "100"), emissao_nova_cessao_texto="01/09/2026")
    credit = replace(_credit("A", "100"), signature_date=date(2026, 8, 15))
    b = prepare_batch(
        PfmiData(payments=[p]), [credit], [_due()],
        liquidation_date=date(2026, 9, 3), first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] == date(2026, 8, 15)
    assert "EMISSAO_NOVA_CESSAO_AUSENTE" not in row["PENDENCIAS"]
    assert "EMISSAO_NOVA_CESSAO_INVALIDA" not in row["PENDENCIAS"]


def test_cessao_emission_date_never_derived_from_signature_or_liquidation():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    p = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
        emissao_nova_cessao_texto="01/09/2026",
    )
    credit = replace(_credit("A", "100"), signature_date=date(2026, 8, 15))
    b = prepare_batch(
        PfmiData(payments=[p]), [credit], [_due()],
        liquidation_date=date(2026, 9, 3), first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] == date(2026, 9, 1)
    assert row["DT_EMISSAO_TITULO"] != credit.signature_date
    assert row["DT_EMISSAO_TITULO"] != date(2026, 9, 3)


def test_conexcred_word_boundary_accepts_prefix_rejects_mid_word_and_middle():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    base = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
    )
    accepted = ("CONEXCRED", "CONEXCRED SERVICOS LTDA", "CONEXCRED INTERMEDIACAO FULANO")
    for name in accepted:
        p = replace(base, beneficiary_name=name)
        b = prepare_batch(
            PfmiData(payments=[p]),
            [_credit("A", "100")],
            [_due()],
            liquidation_date=None,
            first_sequence=None,
        )
        row = b.rows[0].values
        assert row["ESTADO_CNPJ_CONEXCRED"] == "UNICO_VALIDO", name
        assert row["DOC_CEDENTE"] == VALID_CNPJ, name

    # "CONEXCREDITO" starts with the same letters but has no word boundary, and
    # "GRUPO CONEXCRED LTDA" only has the word in the middle: neither counts.
    rejected = ("CONEXCREDITO SERVICOS LTDA", "GRUPO CONEXCRED LTDA")
    for name in rejected:
        p = replace(base, beneficiary_name=name)
        b = prepare_batch(
            PfmiData(payments=[p]),
            [_credit("A", "100")],
            [_due()],
            liquidation_date=None,
            first_sequence=None,
        )
        row = b.rows[0].values
        assert row["ESTADO_CNPJ_CONEXCRED"] == "AUSENTE", name
        assert row["DOC_CEDENTE"] == "", name
        assert "CNPJ_CONEXCRED_AUSENTE" in row["PENDENCIAS"], name


def test_conexcred_cnpj_conflicting_across_multiple_documents():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    p = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("50"),
        beneficiary_value=Decimal("50"),
    )
    other_cnpj = "11444777000161"
    b = prepare_batch(
        PfmiData(payments=[p, replace(p, beneficiary_document=other_cnpj)]),
        [_credit("A", "100")],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    row = b.rows[0].values
    assert row["ESTADO_CNPJ_CONEXCRED"] == "CONFLITANTE"
    assert row["DOC_CEDENTE"] == ""
    assert "CNPJ_CONEXCRED_CONFLITANTE" in row["PENDENCIAS"]


def test_original_credit_invalid_document_blocks_normal_but_not_cessao():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    invalid_doc_credit = replace(_credit("A", "100"), cedent_document="A0002102")

    normal = replace(_payment("B1", "100"))
    b = prepare_batch(
        PfmiData(payments=[normal]),
        [invalid_doc_credit],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    assert "DOC_CEDENTE_INVALIDO" in b.rows[0].values["PENDENCIAS"]

    cessao = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
    )
    b = prepare_batch(
        PfmiData(payments=[cessao]),
        [invalid_doc_credit],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    row = b.rows[0].values
    assert "DOC_CEDENTE_INVALIDO" not in row["PENDENCIAS"]
    assert row["DOC_CEDENTE"] == VALID_CNPJ
    assert row["DOC_CEDENTE_ANALITICO"] == "A0002102"


def test_cessao_never_autofills_emission_date():
    from dataclasses import replace

    from gerador_cnab_nox.models import CESSAO

    credit = replace(_credit("A", "100"), signature_date=date(2026, 8, 1))
    p = replace(
        _payment("B1", "100"),
        modality=CESSAO,
        commission_text="0",
        beneficiary_name="CONEXCRED",
        beneficiary_document=VALID_CNPJ,
        operation_value=Decimal("100"),
        beneficiary_value=Decimal("100"),
    )
    b = prepare_batch(
        PfmiData(payments=[p]),
        [credit],
        [_due()],
        liquidation_date=date(2026, 9, 3),
        first_sequence=1,
    )
    row = b.rows[0].values
    assert row["DT_EMISSAO_TITULO"] is None


def test_t14_lawyer_outside_parentheses_does_not_create_original_identity():
    from dataclasses import replace

    for name in ("FULANO DE TAL (CREDOR SINTETICO)", "BELTRANO (CREDOR SINTETICO)"):
        p = replace(_payment("B1", "100"), failure="FALÊNCIA GAMA", cedent_name=name)
        c = replace(
            _credit("A", "100"),
            campaign="FALENCIA GAMA",
            cedent_name=name,
            cedent_document="123",
        )
        b = prepare_batch(
            PfmiData(payments=[p]),
            [c],
            [],
            liquidation_date=None,
            first_sequence=None,
            lawyer_rules={"FALENCIA GAMA": ("FULANO DE TAL", "BELTRANO")},
        )
        r = b.rows[0].values
        assert r["NOME_CEDENTE"] == name.split(" (")[0]
        assert r["DOC_CEDENTE"] == "123" and r["NOME_CEDENTE_ANALITICO"] == name
        assert r["INCLUIR_CNAB"] == "NAO"
        # Parenthetical creditor alone is no automatic identity match.
        c = replace(c, cedent_name="CREDOR SINTETICO", acquisition_value=Decimal("90"))
        b = prepare_batch(
            PfmiData(payments=[p]),
            [c],
            [],
            liquidation_date=None,
            first_sequence=None,
            lawyer_rules={"FALENCIA GAMA": ("FULANO DE TAL", "BELTRANO")},
        )
        assert b.rows[0].values["ID_CREDITO"] == ""


def test_related_block_suggestion_is_exploratory_and_never_selects():
    from dataclasses import replace

    p1 = replace(_payment("B1", "40"), cedent_name="PESSOA UM")
    p2 = replace(_payment("B2", "60"), cedent_name="TERCEIRO (PESSOA UM)")
    credit = replace(_credit("A", "100"), cedent_name="PESSOA UM")
    b = prepare_batch(
        PfmiData(payments=[p1, p2]),
        [credit],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert len(b.rows) == 2
    assert all(r.values["INCLUIR_CNAB"] == "NAO" for r in b.rows)
    assert len(b.related_suggestions) == 1
    suggestion = b.related_suggestions[0]
    assert suggestion["credor_a"] == "PESSOA UM"
    assert suggestion["credor_b"] == "TERCEIRO (PESSOA UM)"
    assert suggestion["valor_a"] == Decimal("40") and suggestion["valor_b"] == Decimal("60")
    assert suggestion["total"] == Decimal("100")
    assert suggestion["creditos_encontrados"][0]["referencia"] == "A"


def test_related_block_suggestion_coexists_with_aliases_and_lawyer_rules():
    """Regression: with failure aliases and lawyer rules loaded, an unrelated
    baseline group must stay auto-selected while a parentheses-related pair
    only surfaces as an observational suggestion (never auto-approved)."""
    from dataclasses import replace

    baseline_payment = replace(
        _payment("B0", "200"), cedent_name="PESSOA BASE", failure="FALENCIA BETA"
    )
    baseline_credit = replace(
        _credit("BASE", "200"), cedent_name="PESSOA BASE", campaign="FALENCIA BETA"
    )
    p1 = replace(_payment("B1", "40"), cedent_name="PESSOA UM", failure="FALENCIA GAMA ALT")
    p2 = replace(
        _payment("B2", "60"), cedent_name="TERCEIRO (PESSOA UM)", failure="FALENCIA GAMA ALT"
    )
    credit = replace(
        _credit("A", "100"), cedent_name="PESSOA UM", campaign="FALENCIA GAMA"
    )
    b = prepare_batch(
        PfmiData(payments=[baseline_payment, p1, p2]),
        [baseline_credit, credit],
        [_due(), replace(_due(), debtor_name="FALENCIA BETA")],
        liquidation_date=None,
        first_sequence=None,
        failure_aliases={"FALENCIA GAMA ALT": "FALENCIA GAMA"},
        lawyer_rules={"OUTRA FALENCIA": ("ADVOGADO EXEMPLO",)},
    )
    assert len(b.rows) == 3
    by_name = {r.values["NOME_CEDENTE_PFMI"]: r.values for r in b.rows}
    assert by_name["PESSOA BASE"]["STATUS"] == "OK"
    assert by_name["PESSOA BASE"]["INCLUIR_CNAB"] == "SIM"
    assert by_name["PESSOA UM"]["INCLUIR_CNAB"] == "NAO"
    assert by_name["TERCEIRO (PESSOA UM)"]["INCLUIR_CNAB"] == "NAO"
    assert len(b.related_suggestions) == 1
    suggestion = b.related_suggestions[0]
    assert suggestion["total"] == Decimal("100")
    assert suggestion["creditos_encontrados"][0]["referencia"] == "A"


def test_related_block_suggestion_requires_exact_aggregate_match():
    from dataclasses import replace

    p1 = replace(_payment("B1", "40"), cedent_name="PESSOA DOIS")
    p2 = replace(_payment("B2", "61"), cedent_name="TERCEIRO (PESSOA DOIS)")
    credit = replace(_credit("A", "100"), cedent_name="PESSOA DOIS")
    b = prepare_batch(
        PfmiData(payments=[p1, p2]),
        [credit],
        [_due()],
        liquidation_date=None,
        first_sequence=None,
    )
    assert b.related_suggestions == []
