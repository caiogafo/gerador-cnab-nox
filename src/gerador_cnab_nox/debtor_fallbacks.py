"""Read only the delimited data block in the local fallback reference document."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

from .errors import InputFileError
from .normalize import digits, normalize_failure

BEGIN = "<!-- CNAB_NOX_FALLBACKS_BEGIN -->"
END = "<!-- CNAB_NOX_FALLBACKS_END -->"


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"chave duplicada: {key}")
        result[key] = value
    return result


def load_fallbacks(path=None):
    path = Path(path or os.environ.get("CNAB_NOX_EQUIVALENCIAS_PATH") or (
        Path(__file__).resolve().parents[2] / "docs" / "EQUIVALENCIAS_FALENCIAS.md"
    ))
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise InputFileError(f"Não foi possível ler o cadastro de sacados: {path.name}") from exc
    if BEGIN not in text and END not in text:
        return {}
    try:
        if text.count(BEGIN) != 1 or text.count(END) != 1:
            raise ValueError("delimitadores duplicados ou ausentes")
        if text.index(BEGIN) > text.index(END):
            raise ValueError("delimitadores fora de ordem")
        block = text.split(BEGIN, 1)[1].split(END, 1)[0].strip()
        if not block.startswith("```json\n") or not block.endswith("```"):
            raise ValueError("bloco JSON inválido")
        data = json.loads(block[8:-3], object_pairs_hook=_unique_object)
        if not isinstance(data, dict):
            raise ValueError("cadastro deve ser um objeto")
        result = {}
        for failure, entry in data.items():
            raw = entry["cnpj"]
            if not isinstance(raw, str) or not re.fullmatch(r"[0-9. /-]+", raw):
                raise ValueError("CNPJ deve conter somente números e pontuação")
            document = digits(raw)
            if len(document) != 14:
                raise ValueError("CNPJ deve conter 14 dígitos")
            name = entry["nome"]
            key = normalize_failure(failure)
            if not key or key in result or not isinstance(name, str) or not name.strip():
                raise ValueError("falência ou nome inválido/duplicado")
            result[key] = {"document": document, "name": name.strip(), "source": str(path)}
        return result
    except (KeyError, TypeError, ValueError, AttributeError) as exc:
        raise InputFileError(f"Cadastro de sacados inválido: {exc}") from exc
