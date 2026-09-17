"""User-confirmed local failure equivalences; no person identities are stored."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from .errors import InputFileError
from .normalize import normalize_failure, normalize_name


def default_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local" / "share"))
    return base / "GeradorCNABNOX" / "equivalencias_falencias.json"


def validate_aliases(values):
    if not isinstance(values, dict):
        raise InputFileError("Equivalências de falência devem ser uma lista de pares de nomes.")
    normalized = {}
    for source, target in values.items():
        if not isinstance(source, str) or not isinstance(target, str):
            raise InputFileError("Informe nomes de falência em texto.")
        key, value = normalize_name(source), normalize_name(target)
        if not key or not value or key == value:
            raise InputFileError("Informe dois nomes de falência diferentes e não vazios.")
        if key in normalized and normalized[key] != value:
            raise InputFileError("A mesma falência possui destinos conflitantes.")
        normalized[key] = value
    _resolved_aliases(normalized)
    return normalized


def _resolved_aliases(values):
    """Compose local aliases with built-ins without storing canonical self-loops."""
    normalized = {}
    for source, target in values.items():
        key, value = normalize_failure(source), normalize_failure(target)
        if key == value:
            continue
        if key in normalized and normalized[key] != value:
            raise InputFileError("A mesma falência possui destinos conflitantes.")
        normalized[key] = value
    for start in normalized:
        seen = set()
        current = start
        while current in normalized:
            if current in seen:
                raise InputFileError(
                    "Equivalências formam um ciclo. Escolha um nome de referência."
                )
            seen.add(current)
            current = normalized[current]
    return normalized


def load_aliases(path=None):
    path = Path(path) if path is not None else default_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        raise InputFileError("Não foi possível ler as equivalências de falência locais.") from exc
    if not isinstance(data, dict) or data.get("version") != 1:
        raise InputFileError("Formato de equivalências desconhecido. Revise o cadastro local.")
    return validate_aliases(data.get("aliases"))


def save_aliases(values, path=None):
    values = validate_aliases(values)
    path = Path(path) if path is not None else default_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         delete=False) as handle:
            temporary = Path(handle.name)
            json.dump({"version": 1, "aliases": values}, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def resolver(values):
    aliases = _resolved_aliases(validate_aliases(values))

    def resolve(name):
        key = normalize_failure(name)
        while key in aliases:
            key = aliases[key]
        return key

    return resolve
