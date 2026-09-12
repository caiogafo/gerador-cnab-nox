"""Regras locais, offline e determinísticas de identidade advogado/falência.

Sem arquivo local, nenhuma regra é aplicada e o caso cai em revisão manual
(comportamento padrão de matching.py quando lawyer_rules={})."""

from __future__ import annotations

import json
import os
from pathlib import Path

from .errors import InputFileError
from .normalize import normalize_failure, normalize_name


def default_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local" / "share"))
    return base / "GeradorCNABNOX" / "regras_advogado.json"


def validate_rules(data) -> dict[str, tuple[str, ...]]:
    if not isinstance(data, dict) or data.get("version") != 1:
        raise InputFileError("Formato de regras de advogado desconhecido.")
    casos = data.get("casos")
    if not isinstance(casos, list):
        raise InputFileError("Regras de advogado devem ser uma lista de casos.")
    rules: dict[str, tuple[str, ...]] = {}
    for item in casos:
        if not isinstance(item, dict):
            raise InputFileError("Cada regra de advogado deve ser um objeto.")
        failure = item.get("falencia")
        identities = item.get("identidades")
        if not isinstance(failure, str) or not failure.strip():
            raise InputFileError("Informe a falência de cada regra de advogado.")
        if not isinstance(identities, list) or not identities or not all(
            isinstance(i, str) and i.strip() for i in identities
        ):
            raise InputFileError("Informe ao menos uma identidade de advogado por regra.")
        key = normalize_failure(failure)
        rules[key] = tuple(dict.fromkeys(normalize_name(i) for i in identities))
    return rules


def load_rules(path=None) -> dict[str, tuple[str, ...]]:
    path = Path(path) if path is not None else default_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as exc:
        raise InputFileError("Não foi possível ler as regras locais de advogado.") from exc
    return validate_rules(data)
