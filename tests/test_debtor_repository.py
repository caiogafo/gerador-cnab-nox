import json
from dataclasses import replace
from datetime import date

import pytest

from gerador_cnab_nox import debtor_fallbacks as repository
from gerador_cnab_nox.errors import InputFileError
from gerador_cnab_nox.gui_debtors import DebtorEditor
from gerador_cnab_nox.matching import prepare_batch
from gerador_cnab_nox.models import PfmiData

from .conftest import VALID_CNPJ
from .test_matching import _credit, _due, _payment


@pytest.fixture(autouse=True)
def isolated_cache(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))


def markdown(tmp_path):
    path = tmp_path / "fallback.md"
    data = {"BIANCO": {"cnpj": "00000000000100", "nome": "NOME HISTORICO"}}
    path.write_text(f"{repository.BEGIN}\n```json\n{json.dumps(data)}\n```\n"
                    f"{repository.END}\n", encoding="utf-8")
    return path


def test_fallback_merge_priority(tmp_path):
    path = markdown(tmp_path)
    repository.save_debtor(" bianco ", "11.222.333/0001-81", "MASSA FALIDA BIANCO")
    result = repository.load_fallbacks(path)["BIANCO"]
    assert result["document"] == VALID_CNPJ
    assert result["name"] == "MASSA FALIDA BIANCO"
    assert result["source_kind"] == "CACHE_LOCAL"
    assert result["updated_at"].endswith("Z")
    repository.save_debtor("keleti", VALID_CNPJ, "KELETI OFICIAL")
    assert "KELETI ENGENHARIA" in repository.load_local()


def test_gui_rejects_invalid_cnpj(_shared_tk_root):
    saved = []
    editor = DebtorEditor(_shared_tk_root, lambda: saved.append(True))
    try:
        editor.failure.set("BIANCO")
        editor.name.set("MASSA FALIDA BIANCO")
        editor.document.set("11.222.333/0001-82")
        editor.save()
        assert not repository.default_path().exists()
        assert not saved
        assert "CNPJ" in editor.message.get()
        editor.document.set(VALID_CNPJ)
        editor.save()
        assert repository.load_local()["BIANCO"]["document"] == VALID_CNPJ
        assert saved == [True]
        assert len(editor.tree.get_children()) == 1
    finally:
        editor.window.destroy()


def batch(fallbacks, due=None):
    return prepare_batch(
        PfmiData(payments=[replace(_payment("B1", "100"), failure="BIANCO")]),
        [replace(_credit("A", "100"), campaign="BIANCO")], due or [],
        liquidation_date=date(2026, 9, 18), first_sequence=1, debtor_fallbacks=fallbacks,
    )


def test_missing_fallback_generates_pendency(tmp_path):
    row = batch(repository.load_fallbacks(tmp_path / "absent.md")).rows[0].values
    assert "SACADO_NAO_CADASTRADO_NO_SISTEMA" in row["PENDENCIAS"]
    assert row["DOC_SACADO"] == row["NOME_SACADO"] == ""
    assert row["STATUS"] == "PREENCHER_MANUALMENTE"


def test_local_cache_injection_and_valid_base_priority(tmp_path):
    repository.save_debtor("BIANCO", VALID_CNPJ, "MASSA FALIDA BIANCO")
    fallbacks = repository.load_fallbacks(tmp_path / "absent.md")
    row = batch(fallbacks).rows[0].values
    assert row["DOC_SACADO"] == VALID_CNPJ
    assert row["NOME_SACADO"] == "MASSA FALIDA BIANCO"
    assert row["DT_VENCIMENTO"] == date(2026, 9, 18)
    assert "FALLBACK_SACADO_APLICADO_VIA_CACHE_LOCAL" in row["ALERTAS"]
    assert "SACADO_NAO_CADASTRADO_NO_SISTEMA" not in row["PENDENCIAS"]
    row = batch(fallbacks, [replace(_due(), debtor_name="BIANCO")]).rows[0].values
    assert row["DT_VENCIMENTO"] == date(2030, 1, 1)
    assert "FALLBACK_SACADO_APLICADO_VIA_CACHE_LOCAL" not in row["ALERTAS"]


def test_failed_replace_preserves_previous_cache(monkeypatch):
    repository.save_debtor("BIANCO", VALID_CNPJ, "NOME ORIGINAL")
    path = repository.default_path()
    original = path.read_bytes()

    def fail(*args):
        raise OSError("interrupted")

    monkeypatch.setattr(repository.os, "replace", fail)
    with pytest.raises(OSError):
        repository.save_debtor("BIANCO", VALID_CNPJ, "OUTRO NOME")
    assert path.read_bytes() == original
    assert list(path.parent.iterdir()) == [path]


def test_corrupted_cache_is_not_overwritten():
    path = repository.default_path()
    path.parent.mkdir(parents=True)
    path.write_text("{broken", encoding="utf-8")
    with pytest.raises(InputFileError):
        repository.save_debtor("BIANCO", VALID_CNPJ, "NOME")
    assert path.read_text(encoding="utf-8") == "{broken"
