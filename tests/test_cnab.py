from __future__ import annotations

from datetime import date
from decimal import Decimal
from hashlib import sha256
from pathlib import Path

from gerador_cnab_nox.cnab import _scaled_rate, render_cnab
from gerador_cnab_nox.validation import ValidatedBatch

from .conftest import VALID_CNPJ, VALID_CPF


def _row(cedent_document=VALID_CPF, cedent_type=1):
    return {
        "DOC_CEDENTE": cedent_document,
        "NOME_CEDENTE": "ACME LIMITADA",
        "SEU_NUMERO": 123,
        "NU_DOCUMENTO": 123,
        "DT_VENCIMENTO": date(2030, 1, 31),
        "VL_NOMINAL": Decimal("150.00"),
        "DOC_SACADO": VALID_CNPJ,
        "NOME_SACADO": "FALENCIA ALFA SOCIEDADE ANONIMA",
        "VL_PRESENTE": Decimal("100.00"),
        "TIPO_PESSOA_SACADO": 2,
        "ENDERECO": "",
        "CEP": "",
        "TP_TITULO": 24,
        "DT_EMISSAO_TITULO": date(2026, 9, 1),
        "COOBRIGACAO": 2,
        "TIPO_PESSOA_CEDENTE": cedent_type,
        "NFE": "",
        "VALOR_PAGO_TITULO": Decimal("0.00"),
        "INDEXADOR": "",
        "TAXA_INDEXADOR": None,
        "MOVIMENTO": 1,
    }


def test_cnab_structure_and_homologated_cpf_padding() -> None:
    batch = ValidatedBatch([_row()], date(2026, 9, 3), 123, [])
    payload = render_cnab(batch)
    assert not payload.startswith(b"\xef\xbb\xbf")
    assert payload.endswith(b"\r\n")
    records = payload[:-2].split(b"\r\n")
    assert [len(record) for record in records] == [444, 444, 444]
    assert records[0][94:100] == b"030926"
    assert records[1][380:394] == b"   " + VALID_CPF.encode()
    assert records[1][220:234] == VALID_CNPJ.encode()
    assert records[1][438:444] == b"000002"
    assert records[2][438:444] == b"000003"
    golden_hash = (
        (Path(__file__).parent / "fixtures" / "golden_cnab_v1_2.sha256")
        .read_text(encoding="ascii")
        .strip()
    )
    assert sha256(payload).hexdigest() == golden_hash
    from .golden_expected import expected_payload

    assert payload == expected_payload()
    assert payload == (Path(__file__).parent / "fixtures" / "golden_cnab_v1_2.txt").read_bytes()


def test_cnpj_cedent_has_no_padding() -> None:
    batch = ValidatedBatch([_row(VALID_CNPJ, 2)], date(2026, 9, 3), 123, [])
    detail = render_cnab(batch).split(b"\r\n")[1]
    assert detail[380:394] == VALID_CNPJ.encode()


def test_rate_matches_vba_scaled_representation() -> None:
    assert _scaled_rate(Decimal("1")) == "0010000000"
    assert _scaled_rate(Decimal("0.5")) == "0005000000"


def test_t20_comparison_equivalence_does_not_rewrite_output():
    from gerador_cnab_nox.normalize import normalize_name, technical_name

    assert normalize_name("Ácme, Limitada") == normalize_name("ACME LTDA")
    assert technical_name(" Ácme, Limitada ") == "ACME, LIMITADA"
    row = _row()
    row["NOME_CEDENTE"] = "Ácme, Limitada"
    detail = render_cnab(ValidatedBatch([row], date(2026, 9, 3), 123, [])).split(b"\r\n")[1]
    assert detail[334:374] == b"ACME, LIMITADA".ljust(40)


def test_t21_independent_limits_header_trailer_and_cpf_debtor():
    import pytest

    from gerador_cnab_nox.cnab import _trailer
    from gerador_cnab_nox.errors import ValidationError

    row = _row()
    row["VL_NOMINAL"] = Decimal("99999999999.99")
    row["VL_PRESENTE"] = Decimal("99999999999.99")
    row["SEU_NUMERO"] = 9999999999
    row["NU_DOCUMENTO"] = 9999999999
    row["DOC_SACADO"] = VALID_CPF
    row["TIPO_PESSOA_SACADO"] = 1
    detail = render_cnab(ValidatedBatch([row], date(2026, 9, 3), 9999999999, [])).split(b"\r\n")[1]
    assert detail[126:139] == detail[192:205] == b"9999999999999"
    assert detail[110:120] == b"9999999999"
    assert detail[220:234] == b"00052998224725"
    assert _trailer(999999)[438:] == "999999"
    with pytest.raises(ValidationError):
        _trailer(1000000)
    row["VL_NOMINAL"] = Decimal("100000000000.00")
    with pytest.raises(ValidationError):
        render_cnab(ValidatedBatch([row], date(2026, 9, 3), 123, []))


def test_t22_atomic_write_failure_and_no_clobber(tmp_path, monkeypatch):
    import pytest

    import gerador_cnab_nox.cnab as cnab
    from gerador_cnab_nox.errors import ValidationError

    batch = ValidatedBatch([_row()], date(2026, 9, 3), 123, [])
    output = tmp_path / "final.txt"

    def failure(*_args):
        raise PermissionError("synthetic denied")

    with monkeypatch.context() as patch:
        patch.setattr(cnab.os, "link", failure)
        with pytest.raises(PermissionError):
            cnab.write_cnab(batch, output)
    assert not output.exists() and not list(tmp_path.glob(".*.tmp"))
    cnab.write_cnab(batch, output)
    before = output.read_bytes()
    with pytest.raises(ValidationError):
        cnab.write_cnab(batch, output)
    assert output.read_bytes() == before
