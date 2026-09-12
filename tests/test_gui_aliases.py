import os
from tkinter import TclError, Tk

import pytest

from gerador_cnab_nox.failure_aliases import load_aliases
from gerador_cnab_nox.gui_aliases import AliasEditor


def test_editor_requires_confirmation_persists_and_removes(tmp_path):
    try:
        root = Tk()
    except TclError as exc:
        if os.environ.get("CNAB_NOX_REQUIRE_GUI") == "1":
            pytest.fail(str(exc))
        pytest.skip(f"Tk indisponível: {exc}")
    root.withdraw()
    path = tmp_path / "aliases.json"
    changes = []
    try:
        editor = AliasEditor(root, lambda: changes.append(True), path)
        editor.source.set("ALFA")
        editor.target.set("MASSA ALFA")
        editor.save()
        assert not path.exists()
        editor.confirmed.set(True)
        editor.save()
        assert load_aliases(path) == {"ALFA": "MASSA ALFA"}
        assert changes == [True]
        editor.confirmed.set(True)
        editor.target.set("OUTRO")
        assert not editor.confirmed.get()
        editor.tree.selection_set(editor.tree.get_children()[0])
        editor.remove()
        assert load_aliases(path) == {}
        assert len(changes) == 2
    finally:
        root.destroy()
