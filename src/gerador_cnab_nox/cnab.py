from __future__ import annotations

import os
import tempfile
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

from .errors import ValidationError
from .normalize import money, technical_name
from .validation import ValidatedBatch

ORIGINATOR_CODE = "00000000000000000125"


def render_cnab(batch: ValidatedBatch) -> bytes:
    records = [_header(batch.liquidation_date)]
    for physical_sequence, row in enumerate(batch.rows, start=2):
        records.append(_detail(row, batch.liquidation_date, physical_sequence))
    records.append(_trailer(len(records) + 1))
    issues: list[str] = []
    encoded: list[bytes] = []
    for index, record in enumerate(records, start=1):
        try:
            payload = record.encode("cp1252")
        except UnicodeEncodeError:
            issues.append(f"Registro {index} incompatível com Windows-1252")
            continue
        if len(payload) != 444:
            issues.append(f"Registro {index} possui {len(payload)} bytes, esperado 444")
        encoded.append(payload)
    if issues:
        raise ValidationError(issues)
    return b"\r\n".join(encoded) + b"\r\n"


def write_cnab(batch: ValidatedBatch, output_path: str | Path) -> Path:
    destination = Path(output_path)
    if destination.suffix.lower() != ".txt":
        raise ValidationError(["A saída CNAB deve usar a extensão .txt"])
    if destination.exists():
        raise ValidationError([f"A saída já existe e não será sobrescrita: {destination.name}"])
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload = render_cnab(batch)
    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{destination.stem}-", suffix=".tmp", dir=destination.parent
    )
    os.close(descriptor)
    temporary = Path(temp_name)
    try:
        temporary.write_bytes(payload)
        os.link(temporary, destination)  # atomic no-clobber publication
    finally:
        temporary.unlink(missing_ok=True)
    return destination


def _header(liquidation_date: date) -> str:
    record = (
        "0"
        + "1"
        + "REMESSA"
        + "01"
        + "COBRANCA".ljust(15)
        + ORIGINATOR_CODE
        + "0" * 30
        + "001"
        + "0" * 15
        + _date6(liquidation_date)
        + " " * 8
        + "MX"
        + "0000001"
        + " " * 321
        + "000001"
    )
    return record


def _detail(row: dict[str, Any], liquidation_date: date, physical_sequence: int) -> str:
    cedent_type = int(row["TIPO_PESSOA_CEDENTE"])
    debtor_type = int(row["TIPO_PESSOA_SACADO"])
    cedent_doc = str(row["DOC_CEDENTE"])
    debtor_doc = str(row["DOC_SACADO"])
    rate = row.get("TAXA_INDEXADOR")
    rate_field = " " * 10 if rate is None else _scaled_rate(Decimal(rate))
    record = (
        "1"
        + " " * 6
        + _left(str(row.get("INDEXADOR") or " "), 1, " ")
        + " " * 2
        + rate_field
        + _numeric(row["COOBRIGACAO"], 2)
        + "0" * 15
        + _right(str(row["SEU_NUMERO"]), 25, " ")
        + "001"
        + "0" * 5
        + "0" * 11
        + "1"
        + _amount(row["VALOR_PAGO_TITULO"], 10)
        + "1"
        + "N"
        + _date6(liquidation_date)
        + " " * 4
        + " "
        + "1"
        + " " * 2
        + _numeric(row["MOVIMENTO"], 2)
        + _right(str(row["NU_DOCUMENTO"]), 10, " ")
        + _date6(row["DT_VENCIMENTO"])
        + _amount(row["VL_NOMINAL"], 13)
        + "0" * 3
        + "0" * 5
        + _numeric(row["TP_TITULO"], 2)
        + " "
        + _date6(row["DT_EMISSAO_TITULO"])
        + "0" * 3
        + _numeric(cedent_type, 2)
        + "0" * 12
        + "0" * 6
        + "0" * 13
        + _amount(row["VL_PRESENTE"], 13)
        + "0" * 13
        + _numeric(debtor_type, 2)
        + (("000" + debtor_doc) if debtor_type == 1 else debtor_doc)
        + _left(technical_name(row["NOME_SACADO"])[:40], 40, " ")
        + _left(str(row.get("ENDERECO") or ""), 40, " ")
        + " " * 12
        + _left(str(row.get("CEP") or ""), 8, "0")
        + _left(technical_name(row["NOME_CEDENTE"])[:40], 40, " ")
        + " " * 6
        + ((" " * 3 + cedent_doc) if cedent_type == 1 else cedent_doc)
        + _right(str(row.get("NFE") or ""), 44, "0")
        + _numeric(physical_sequence, 6)
    )
    return record


def _trailer(total_records: int) -> str:
    return "9" + " " * 437 + _numeric(total_records, 6)


def _date6(value: date) -> str:
    return value.strftime("%d%m%y")


def _amount(value: Decimal, width: int) -> str:
    try:
        decimal_value = money(value)
    except ValueError as exc:
        raise ValidationError(["Valor monetário fora do leiaute"]) from exc
    cents = int(decimal_value * 100)
    return _numeric(cents, width)


def _scaled_rate(value: Decimal) -> str:
    scaled = int(value.quantize(Decimal("0.0000001"), rounding=ROUND_HALF_UP) * 10_000_000)
    return _numeric(scaled, 10)


def _numeric(value: Any, width: int) -> str:
    text = str(int(value))
    if len(text) > width or not text.isdigit():
        raise ValidationError([f"Valor {value!r} não cabe em campo numérico de {width} posições"])
    return text.rjust(width, "0")


def _left(value: str, width: int, fill: str) -> str:
    if len(value) > width:
        raise ValidationError([f"Texto excede campo de {width} posições"])
    return value.ljust(width, fill)


def _right(value: str, width: int, fill: str) -> str:
    if len(value) > width:
        raise ValidationError([f"Texto excede campo de {width} posições"])
    return value.rjust(width, fill)
