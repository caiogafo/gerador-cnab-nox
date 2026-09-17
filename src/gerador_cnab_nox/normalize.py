from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation, localcontext
from typing import Any

CREDOR_ALIASES = {
    # Mapeamento do Agenor (ajuste o valor da direita para o nome exato no Analítico)
    "AGENOR LUIZ DE SOUZA FILHO": "AGENOR LUIZ DE SOUZA", 
    "MARIA ANGELICA NALIN 2": "MARIA ANGELICA NALIN",
}

CENT = Decimal("0.01")
MAX_MONEY = Decimal("99999999999.99")
# Aggregate source values do not occupy an individual CNAB monetary field.
# 26 decimal positions, including cents, stay exact in Decimal's default context.
MAX_AGGREGATE = Decimal("1e24")
FORMULA_PREFIXES = ("=", "+", "-", "@")


def header_key(value: Any) -> str:
    text = strip_accents(str(value or "")).upper()
    return re.sub(r"[^A-Z0-9]+", "", text)


def strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(char for char in normalized if not unicodedata.combining(char))

def normalize_name(value: Any) -> str:
    # 1. Padronização primária
    text = strip_accents(str(value or "")).upper()
    
    # 2. Limpeza de prefixos jurídicos (Espólios)
    padrao_prefixos = r'^(ESPOLIO(?:\s+DE)?\.?|INVENTARIANTE(?:\s+DE)?\.?|ESP\.?\s+DE)\s+'
    text = re.sub(padrao_prefixos, "", text)

    # 3. Substituições corporativas contratuais
    for pattern, replacement in (
        (r"(?<![A-Z0-9])L\s*\.\s*T\s*\.\s*D\s*\.\s*A\s*\.?(?![A-Z0-9])", "LTDA"),
        (r"(?<![A-Z0-9])C\s*\.\s*I\s*\.\s*A\s*\.?(?![A-Z0-9])", "CIA"),
        (r"(?<![A-Z0-9])S\s*[./]\s*A\s*\.?(?![A-Z0-9])", "SA"),
    ):
        text = re.sub(pattern, f" {replacement} ", text)

    # 4. Remoção de caracteres especiais
    text = re.sub(r"[^A-Z0-9]+", " ", text).strip()

    # 5. Tratamento de tokens
    tokens = text.split()
    canonical: list[str] = []
    index = 0
    while index < len(tokens):
        pair = " ".join(tokens[index : index + 2])
        if pair == "SOCIEDADE ANONIMA":
            canonical.append("SA")
            index += 2
            continue
        token = tokens[index]
        canonical.append({"LIMITADA": "LTDA", "COMPANHIA": "CIA"}.get(token, token))
        index += 1

    return " ".join(canonical)


_NAME_JOIN_STOPWORDS = frozenset({"DA", "DE", "DO", "DAS", "DOS", "E"})
_VOWELS = frozenset("AEIOU")


def _fold_sz(token: str) -> str:
    """Fold S/Z at a word's final or intervocalic position to one spelling.

    Only that narrow, documented Portuguese variant class (LUIZ/LUIS,
    SOUZA/SOUSA) is folded - never a bare similarity score, and only ever
    used together with an independent, checksum-confirmed document match
    (see matching.py); it never substitutes for identity confirmation on
    its own.
    """
    chars = list(token)
    last = len(chars) - 1
    for index, char in enumerate(chars):
        if char not in ("S", "Z"):
            continue
        at_end = index == last
        intervocalic = (
            0 < index < last and chars[index - 1] in _VOWELS and chars[index + 1] in _VOWELS
        )
        if at_end or intervocalic:
            chars[index] = "S"
    return "".join(chars)


def names_equivalent_under_confirmed_document(name_a: Any, name_b: Any) -> bool:
    """Deterministic name-spelling equivalence, gated by the CALLER already
    having confirmed both names refer to the same document (never called to
    decide identity by itself). Ignores common Portuguese connective
    stopwords and folds the documented S/Z variant class only; never a
    probabilistic/fuzzy similarity score."""

    def fold(value: Any) -> tuple[str, ...]:
        tokens = normalize_name(value).split()
        return tuple(_fold_sz(token) for token in tokens if token not in _NAME_JOIN_STOPWORDS)

    return fold(name_a) == fold(name_b)


def normalize_name_whitespace(value: Any) -> str:
    text = "".join(
        " " if char.isspace() else char
        for char in str(value or "")
    )
    return " ".join(text.split())


def technical_name(value: Any) -> str:
    text = strip_accents(str(value or "")).upper()
    if any(unicodedata.category(char).startswith("C") for char in text):
        raise ValueError("nome possui caractere de controle/incompatível")
    text = " ".join(text.split())
    if not any(char.isalnum() for char in text):
        raise ValueError("nome ausente ou sem caracteres alfanuméricos")
    try:
        text.encode("cp1252")
    except UnicodeEncodeError as exc:
        raise ValueError("nome incompatível com Windows-1252") from exc
    return text

def normalize_failure(value: Any) -> str:
    name = normalize_name(value)
    return {
        "MOGIANO": "MOGIANO TRANSP GERAIS",
        "KELETI": "KELETI ENGENHARIA",
    }.get(name, name)

def commission_rate(value: Any, modality: str) -> Decimal:
    if modality == "NORMAL":
        return Decimal(0)
    if modality != "CESSAO_DA_CESSAO":
        raise ValueError("MODALIDADE_INVALIDA")
    text = str(value if value is not None else "").strip()
    if not re.fullmatch(r"\d+(?:[.,]\d+)?", text, flags=re.ASCII):
        raise ValueError("COMISSAO_PARAMETRO_INVALIDO")
    integer, _, fraction = text.replace(",", ".").partition(".")
    significant = (integer + fraction).lstrip("0") or "0"
    if len(integer.lstrip("0") or "0") > 28 or len(fraction) > 28 or len(significant) > 28:
        raise ValueError("COMISSAO_PARAMETRO_INVALIDO")
    return Decimal(text.replace(",", "."))


def calculate_commission(nominal: Any, acquisition: Any, rate: Decimal):
    n, a = money(nominal), money(acquisition)
    if n is None or a is None or n < 0 or a < 0:
        raise ValueError("VALOR_CANDIDATO_INVALIDO")
    with localcontext() as context:
        context.prec = 80
        base = n - a
        if base < 0 and rate > 0:
            raise ValueError("COMISSAO_BASE_NEGATIVA")
        commission = money(base * rate / 100)
        expected = money(a + commission)
    return base, commission, expected


def digits(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\D", "", str(value))


def mask_document(value: Any) -> str:
    # Only for human review lists (e.g. composition manual-selection
    # candidates), never for the CNAB output itself. Keeps just enough of the
    # digits to distinguish entries without exposing the full CPF/CNPJ.
    doc = digits(value)
    if len(doc) <= 4:
        return "*" * len(doc)
    return doc[:2] + "*" * (len(doc) - 4) + doc[-2:]


def validate_document(value: Any) -> tuple[str, int | None]:
    doc = digits(value)
    if isinstance(value, str) and re.search(r"[^0-9./\-\s]", value):
        return doc, None
    if len(doc) == 11 and _valid_cpf(doc):
        return doc, 1
    if len(doc) == 14 and _valid_cnpj(doc):
        return doc, 2
    return doc, None


def _valid_cpf(doc: str) -> bool:
    if len(set(doc)) == 1:
        return False
    numbers = [int(char) for char in doc]
    for position in (9, 10):
        total = sum(numbers[index] * (position + 1 - index) for index in range(position))
        check = (total * 10) % 11
        if check == 10:
            check = 0
        if numbers[position] != check:
            return False
    return True


def _valid_cnpj(doc: str) -> bool:
    if len(set(doc)) == 1:
        return False
    numbers = [int(char) for char in doc]
    for size, weights in (
        (12, (5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)),
        (13, (6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)),
    ):
        total = sum(numbers[index] * weights[index] for index in range(size))
        remainder = total % 11
        check = 0 if remainder < 2 else 11 - remainder
        if numbers[size] != check:
            return False
    return True


def money(value: Any, *, allow_blank: bool = False, maximum: Decimal = MAX_MONEY) -> Decimal | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        if allow_blank:
            return None
        raise ValueError("valor ausente")
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, (int, float)):
        try:
            result = Decimal(str(value))
        except InvalidOperation as exc:
            raise ValueError("valor inválido") from exc
    else:
        text = str(value).strip().replace("R$", "").replace(" ", "")
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        try:
            result = Decimal(text)
        except InvalidOperation as exc:
            raise ValueError("valor inválido") from exc
    if not result.is_finite():
        raise ValueError("valor não finito")
    if abs(result) > maximum + CENT:
        raise ValueError("valor fora da faixa suportada")
    try:
        with localcontext() as context:
            context.prec = 80
            result = result.quantize(CENT, rounding=ROUND_HALF_UP)
        if abs(result) > maximum:
            raise ValueError("valor fora da faixa suportada")
        return result
    except InvalidOperation as exc:
        raise ValueError("valor fora da faixa suportada") from exc


def parse_date(value: Any, *, allow_blank: bool = False) -> date | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        if allow_blank:
            return None
        raise ValueError("data ausente")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Datas seriais Excel: 1899-12-30 é a base compatível com openpyxl.
        from datetime import timedelta

        try:
            return date(1899, 12, 30) + timedelta(days=int(value))
        except (OverflowError, ValueError) as exc:
            raise ValueError(f"data inválida: {value}") from exc
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"data inválida: {value}")


def parse_positive_int(value: Any, *, allow_blank: bool = False) -> int | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        if allow_blank:
            return None
        raise ValueError("número ausente")
    if isinstance(value, bool):
        raise ValueError("número inválido")
    if isinstance(value, (float, Decimal)) and (
        not Decimal(str(value)).is_finite()
        or Decimal(str(value)) != Decimal(str(value)).to_integral_value()
    ):
        raise ValueError("número deve ser inteiro")
    try:
        number = int(str(value).strip()) if isinstance(value, str) else int(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError("número inválido") from exc
    if number <= 0:
        raise ValueError("número deve ser positivo")
    return number


def excel_safe(value: Any) -> Any:
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        return "'" + value
    return value


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")
