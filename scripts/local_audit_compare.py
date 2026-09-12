"""Independent, privacy-safe parser and byte comparator for NOX CNAB 444 files.

This module intentionally imports no application code. ``parse_cnab`` validates the
structural envelope (CP1252, 444-byte records, CRLF including the final delimiter,
record kinds and the trailer count) and returns the raw records plus the 21 logical
detail fields. Dates are ``date`` objects and monetary values/rates are ``Decimal``.

``compare_cnab`` never raises for malformed CNAB payloads. It returns a JSON-serializable
report containing validation diagnostics, whole-payload hashes/counts and record/field
differences. Difference entries expose only positions, counts and SHA-256 hashes; they
never include field contents.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from difflib import SequenceMatcher
from typing import Any

RECORD_LENGTH = 444
CRLF = b"\r\n"

# Inclusive, one-based ranges for the 21 final fields described by Specs 1.1 section 6.
# The six bytes after NOME_CEDENTE are layout padding, not part of the logical name.
FIELD_RANGES: dict[str, tuple[int, int]] = {
    "DOC_CEDENTE": (381, 394),
    "NOME_CEDENTE": (335, 374),
    "SEU_NUMERO": (38, 62),
    "NU_DOCUMENTO": (111, 120),
    "DT_VENCIMENTO": (121, 126),
    "VL_NOMINAL": (127, 139),
    "DOC_SACADO": (221, 234),
    "NOME_SACADO": (235, 274),
    "VL_PRESENTE": (193, 205),
    "TIPO_PESSOA_SACADO": (219, 220),
    "ENDERECO": (275, 314),
    "CEP": (327, 334),
    "TP_TITULO": (148, 149),
    "DT_EMISSAO_TITULO": (151, 156),
    "COOBRIGACAO": (21, 22),
    "TIPO_PESSOA_CEDENTE": (160, 161),
    "NFE": (395, 438),
    "VALOR_PAGO_TITULO": (83, 92),
    "INDEXADOR": (8, 8),
    "TAXA_INDEXADOR": (11, 20),
    "MOVIMENTO": (109, 110),
}

_HEADER_LAYOUT = [
    ("TIPO_REGISTRO", 1, 1),
    ("HEADER_IDENTIFICACAO_ARQUIVO", 2, 2),
    ("HEADER_LITERAL_REMESSA", 3, 9),
    ("HEADER_CODIGO_SERVICO", 10, 11),
    ("HEADER_LITERAL_SERVICO", 12, 26),
    ("HEADER_ORIGINADOR", 27, 46),
    ("HEADER_RESERVADO_ZEROS_01", 47, 76),
    ("HEADER_BANCO", 77, 79),
    ("HEADER_RESERVADO_ZEROS_02", 80, 94),
    ("DATA_LIQUIDACAO", 95, 100),
    ("HEADER_PADDING_01", 101, 108),
    ("HEADER_IDENTIFICADOR", 109, 110),
    ("HEADER_NUMERO_REMESSA", 111, 117),
    ("HEADER_PADDING_02", 118, 438),
    ("SEQUENCIA_FISICA", 439, 444),
]

_DETAIL_LAYOUT = [
    ("TIPO_REGISTRO", 1, 1),
    ("PADDING_001", 2, 7),
    ("INDEXADOR", 8, 8),
    ("PADDING_002", 9, 10),
    ("TAXA_INDEXADOR", 11, 20),
    ("COOBRIGACAO", 21, 22),
    ("RESERVADO_ZEROS_001", 23, 37),
    ("SEU_NUMERO", 38, 62),
    ("CONSTANTE_BANCO", 63, 65),
    ("RESERVADO_ZEROS_002", 66, 70),
    ("RESERVADO_ZEROS_003", 71, 81),
    ("CONSTANTE_CARTEIRA", 82, 82),
    ("VALOR_PAGO_TITULO", 83, 92),
    ("CONSTANTE_ACEITE", 93, 93),
    ("CONSTANTE_ESPECIE", 94, 94),
    ("DATA_LIQUIDACAO", 95, 100),
    ("PADDING_003", 101, 104),
    ("PADDING_004", 105, 105),
    ("CONSTANTE_INSTRUCAO", 106, 106),
    ("PADDING_005", 107, 108),
    ("MOVIMENTO", 109, 110),
    ("NU_DOCUMENTO", 111, 120),
    ("DT_VENCIMENTO", 121, 126),
    ("VL_NOMINAL", 127, 139),
    ("RESERVADO_ZEROS_004", 140, 142),
    ("RESERVADO_ZEROS_005", 143, 147),
    ("TP_TITULO", 148, 149),
    ("PADDING_006", 150, 150),
    ("DT_EMISSAO_TITULO", 151, 156),
    ("RESERVADO_ZEROS_006", 157, 159),
    ("TIPO_PESSOA_CEDENTE", 160, 161),
    ("RESERVADO_ZEROS_007", 162, 192),
    ("VL_PRESENTE", 193, 205),
    ("RESERVADO_ZEROS_008", 206, 218),
    ("TIPO_PESSOA_SACADO", 219, 220),
    ("DOC_SACADO", 221, 234),
    ("NOME_SACADO", 235, 274),
    ("ENDERECO", 275, 314),
    ("PADDING_007", 315, 326),
    ("CEP", 327, 334),
    ("NOME_CEDENTE", 335, 374),
    ("PADDING_NOME_CEDENTE", 375, 380),
    ("DOC_CEDENTE", 381, 394),
    ("NFE", 395, 438),
    ("SEQUENCIA_FISICA", 439, 444),
]

_TRAILER_LAYOUT = [
    ("TIPO_REGISTRO", 1, 1),
    ("TRAILER_PADDING", 2, 438),
    ("TRAILER_QUANTIDADE_REGISTROS", 439, 444),
]

_UNKNOWN_LAYOUT = [
    ("TIPO_REGISTRO", 1, 1),
    ("POSICOES_DESCONHECIDAS", 2, 444),
]


class CNABParseError(ValueError):
    """Structural/parser errors that contain no decoded CNAB field values."""

    def __init__(self, errors: list[str]):
        self.errors = tuple(errors)
        super().__init__("; ".join(errors))


def _sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _field(record: bytes, start: int, end: int) -> bytes:
    return record[start - 1 : end]


def _raw_records(payload: bytes) -> list[bytes]:
    body = payload[:-2] if payload.endswith(CRLF) else payload
    return body.split(CRLF) if body else []


def _parse_int(raw: bytes, label: str, record_number: int, errors: list[str]) -> int | None:
    stripped = raw.strip(b" ")
    if not stripped or not stripped.isdigit():
        errors.append(f"registro {record_number}: {label} deve conter somente algarismos e padding")
        return None
    return int(stripped)


def _parse_amount(raw: bytes, label: str, record_number: int, errors: list[str]) -> Decimal | None:
    value = _parse_int(raw, label, record_number, errors)
    return None if value is None else Decimal(value) / Decimal(100)


def _parse_date(raw: bytes, label: str, record_number: int, errors: list[str]) -> date | None:
    if len(raw) != 6 or not raw.isdigit():
        errors.append(f"registro {record_number}: {label} deve usar ddmmaa em 6 algarismos")
        return None
    try:
        return datetime.strptime(raw.decode("ascii"), "%d%m%y").date()
    except ValueError:
        errors.append(f"registro {record_number}: {label} contém data civil inválida")
        return None


def _decode_text(raw: bytes) -> str:
    return raw.decode("cp1252").rstrip(" ")


def _parse_document(
    raw: bytes,
    person_type: int | None,
    label: str,
    record_number: int,
    errors: list[str],
    *,
    cpf_prefix: bytes,
) -> str:
    if person_type == 1:
        if not raw.startswith(cpf_prefix) or not raw[len(cpf_prefix) :].isdigit():
            errors.append(
                f"registro {record_number}: {label} CPF não respeita prefixo e 11 algarismos"
            )
            return ""
        return raw[len(cpf_prefix) :].decode("ascii")
    if person_type == 2:
        if len(raw) != 14 or not raw.isdigit():
            errors.append(f"registro {record_number}: {label} CNPJ deve ter 14 algarismos")
            return ""
        return raw.decode("ascii")
    errors.append(f"registro {record_number}: tipo de pessoa de {label} deve ser 01 ou 02")
    return ""


def _parse_detail(record: bytes, record_number: int, errors: list[str]) -> dict[str, Any]:
    cedent_type = _parse_int(
        _field(record, *FIELD_RANGES["TIPO_PESSOA_CEDENTE"]),
        "TIPO_PESSOA_CEDENTE",
        record_number,
        errors,
    )
    debtor_type = _parse_int(
        _field(record, *FIELD_RANGES["TIPO_PESSOA_SACADO"]),
        "TIPO_PESSOA_SACADO",
        record_number,
        errors,
    )
    rate_raw = _field(record, *FIELD_RANGES["TAXA_INDEXADOR"])
    if rate_raw == b" " * len(rate_raw):
        rate: Decimal | None = None
    else:
        scaled_rate = _parse_int(rate_raw, "TAXA_INDEXADOR", record_number, errors)
        rate = None if scaled_rate is None else Decimal(scaled_rate) / Decimal(10_000_000)

    detail: dict[str, Any] = {
        "DOC_CEDENTE": _parse_document(
            _field(record, *FIELD_RANGES["DOC_CEDENTE"]),
            cedent_type,
            "DOC_CEDENTE",
            record_number,
            errors,
            cpf_prefix=b"   ",
        ),
        "NOME_CEDENTE": _decode_text(_field(record, *FIELD_RANGES["NOME_CEDENTE"])),
        "SEU_NUMERO": _parse_int(
            _field(record, *FIELD_RANGES["SEU_NUMERO"]),
            "SEU_NUMERO",
            record_number,
            errors,
        ),
        "NU_DOCUMENTO": _parse_int(
            _field(record, *FIELD_RANGES["NU_DOCUMENTO"]),
            "NU_DOCUMENTO",
            record_number,
            errors,
        ),
        "DT_VENCIMENTO": _parse_date(
            _field(record, *FIELD_RANGES["DT_VENCIMENTO"]),
            "DT_VENCIMENTO",
            record_number,
            errors,
        ),
        "VL_NOMINAL": _parse_amount(
            _field(record, *FIELD_RANGES["VL_NOMINAL"]),
            "VL_NOMINAL",
            record_number,
            errors,
        ),
        "DOC_SACADO": _parse_document(
            _field(record, *FIELD_RANGES["DOC_SACADO"]),
            debtor_type,
            "DOC_SACADO",
            record_number,
            errors,
            cpf_prefix=b"000",
        ),
        "NOME_SACADO": _decode_text(_field(record, *FIELD_RANGES["NOME_SACADO"])),
        "VL_PRESENTE": _parse_amount(
            _field(record, *FIELD_RANGES["VL_PRESENTE"]),
            "VL_PRESENTE",
            record_number,
            errors,
        ),
        "TIPO_PESSOA_SACADO": debtor_type,
        "ENDERECO": _decode_text(_field(record, *FIELD_RANGES["ENDERECO"])),
        "CEP": _decode_text(_field(record, *FIELD_RANGES["CEP"])),
        "TP_TITULO": _parse_int(
            _field(record, *FIELD_RANGES["TP_TITULO"]),
            "TP_TITULO",
            record_number,
            errors,
        ),
        "DT_EMISSAO_TITULO": _parse_date(
            _field(record, *FIELD_RANGES["DT_EMISSAO_TITULO"]),
            "DT_EMISSAO_TITULO",
            record_number,
            errors,
        ),
        "COOBRIGACAO": _parse_int(
            _field(record, *FIELD_RANGES["COOBRIGACAO"]),
            "COOBRIGACAO",
            record_number,
            errors,
        ),
        "TIPO_PESSOA_CEDENTE": cedent_type,
        "NFE": _decode_text(_field(record, *FIELD_RANGES["NFE"])),
        "VALOR_PAGO_TITULO": _parse_amount(
            _field(record, *FIELD_RANGES["VALOR_PAGO_TITULO"]),
            "VALOR_PAGO_TITULO",
            record_number,
            errors,
        ),
        "INDEXADOR": _decode_text(_field(record, *FIELD_RANGES["INDEXADOR"])),
        "TAXA_INDEXADOR": rate,
        "MOVIMENTO": _parse_int(
            _field(record, *FIELD_RANGES["MOVIMENTO"]),
            "MOVIMENTO",
            record_number,
            errors,
        ),
    }
    return detail


def parse_cnab(payload: bytes) -> dict[str, Any]:
    """Parse a strict CNAB 444 payload without importing the production renderer.

    Returns ``records`` (header/details/trailer as bytes), ``details`` (only detail
    records, each with exactly the 21 final fields), ``liquidation_date`` (``date``)
    and ``first_sequence`` (``int``). Raises ``CNABParseError`` with privacy-safe
    diagnostics when the structural envelope or a typed logical field is invalid.
    """

    if not isinstance(payload, bytes):
        raise TypeError("payload deve ser bytes")

    errors: list[str] = []
    if payload.startswith((b"\xef\xbb\xbf", b"\xff\xfe", b"\xfe\xff")):
        errors.append("arquivo: BOM não permitido")
    if not payload.endswith(CRLF):
        errors.append("arquivo: deve terminar com CRLF")

    records = _raw_records(payload)
    if not records:
        errors.append("arquivo: nenhum registro encontrado")
    all_records_decodable = True
    for number, record in enumerate(records, start=1):
        if b"\r" in record or b"\n" in record:
            errors.append(f"registro {number}: quebra de linha fora do padrão CRLF")
        if len(record) != RECORD_LENGTH:
            errors.append(
                f"registro {number}: possui {len(record)} bytes; esperado {RECORD_LENGTH}"
            )
            continue
        try:
            decoded = record.decode("cp1252")
        except UnicodeDecodeError as exc:
            all_records_decodable = False
            errors.append(
                f"registro {number}: byte incompatível com CP1252 na posição {exc.start + 1}"
            )
            continue
        control_position = next(
            (index for index, char in enumerate(decoded, start=1) if ord(char) < 32), None
        )
        if control_position is not None:
            errors.append(f"registro {number}: caractere de controle na posição {control_position}")

    complete_records = [record for record in records if len(record) == RECORD_LENGTH]
    if len(records) < 3:
        errors.append("arquivo: esperado header, ao menos um detalhe e trailer")
    if records and len(records[0]) == RECORD_LENGTH and records[0][:1] != b"0":
        errors.append("registro 1: tipo esperado 0 (header)")
    if records and len(records[-1]) == RECORD_LENGTH and records[-1][:1] != b"9":
        errors.append(f"registro {len(records)}: tipo esperado 9 (trailer)")
    for number, record in enumerate(records[1:-1], start=2):
        if len(record) == RECORD_LENGTH and record[:1] != b"1":
            errors.append(f"registro {number}: tipo esperado 1 (detalhe)")

    liquidation_date: date | None = None
    details: list[dict[str, Any]] = []
    if len(complete_records) == len(records) and records and all_records_decodable:
        liquidation_date = _parse_date(_field(records[0], 95, 100), "DATA_LIQUIDACAO", 1, errors)
        for number, record in enumerate(records[1:-1], start=2):
            detail_liquidation = _parse_date(
                _field(record, 95, 100), "DATA_LIQUIDACAO", number, errors
            )
            if (
                liquidation_date is not None
                and detail_liquidation is not None
                and detail_liquidation != liquidation_date
            ):
                errors.append(f"registro {number}: DATA_LIQUIDACAO difere da data do header")
            details.append(_parse_detail(record, number, errors))

        trailer_count = _parse_int(
            _field(records[-1], 439, 444),
            "TRAILER_QUANTIDADE_REGISTROS",
            len(records),
            errors,
        )
        if trailer_count is not None and trailer_count != len(records):
            errors.append("trailer: quantidade de registros diverge do arquivo")

    if errors:
        raise CNABParseError(errors)
    assert liquidation_date is not None
    assert details
    first_sequence = details[0]["SEU_NUMERO"]
    assert isinstance(first_sequence, int)
    return {
        "records": records,
        "details": details,
        "liquidation_date": liquidation_date,
        "first_sequence": first_sequence,
    }


def _byte_difference_count(actual: bytes, expected: bytes) -> int:
    shared = min(len(actual), len(expected))
    return sum(actual[index] != expected[index] for index in range(shared)) + abs(
        len(actual) - len(expected)
    )


def _record_layout(record: bytes) -> list[tuple[str, int, int]]:
    record_type = record[:1]
    if record_type == b"0":
        return _HEADER_LAYOUT
    if record_type == b"1":
        return _DETAIL_LAYOUT
    if record_type == b"9":
        return _TRAILER_LAYOUT
    return _UNKNOWN_LAYOUT


def _difference_intervals(actual: bytes, expected: bytes, start: int, end: int) -> list[dict]:
    intervals: list[dict[str, int]] = []
    open_start: int | None = None
    for position in range(start, end + 1):
        offset = position - 1
        different = offset >= len(actual) or offset >= len(expected)
        if not different:
            different = actual[offset] != expected[offset]
        if different and open_start is None:
            open_start = position
        if not different and open_start is not None:
            intervals.append({"start": open_start, "end": position - 1})
            open_start = None
    if open_start is not None:
        intervals.append({"start": open_start, "end": end})
    return intervals


def _field_differences(actual: bytes, expected: bytes) -> tuple[list[dict], list[dict]]:
    layout_record = expected if expected else actual
    field_differences: list[dict[str, Any]] = []
    all_intervals: list[dict[str, Any]] = []
    for name, start, end in _record_layout(layout_record):
        actual_field = actual[start - 1 : end]
        expected_field = expected[start - 1 : end]
        if actual_field == expected_field:
            continue
        intervals = _difference_intervals(actual, expected, start, end)
        byte_count = _byte_difference_count(actual_field, expected_field)
        field_differences.append(
            {
                "field": name,
                "start": start,
                "end": end,
                "byte_difference_count": byte_count,
                "intervals": intervals,
                "actual_hash": _sha256(actual_field),
                "expected_hash": _sha256(expected_field),
            }
        )
        all_intervals.extend({"field": name, **interval} for interval in intervals)
    return field_differences, all_intervals


def _changed_record(
    actual: bytes,
    expected: bytes,
    actual_number: int,
    expected_number: int,
) -> dict[str, Any]:
    fields, intervals = _field_differences(actual, expected)
    return {
        "status": "different",
        "record": actual_number,
        "actual_record": actual_number,
        "expected_record": expected_number,
        "actual_type": actual[:1].decode("latin1") if actual else None,
        "expected_type": expected[:1].decode("latin1") if expected else None,
        "byte_difference_count": _byte_difference_count(actual, expected),
        "actual_hash": _sha256(actual),
        "expected_hash": _sha256(expected),
        "fields": fields,
        "intervals": intervals,
    }


def _missing_record(expected: bytes, expected_number: int) -> dict[str, Any]:
    return {
        "status": "missing",
        "record": expected_number,
        "actual_record": None,
        "expected_record": expected_number,
        "actual_type": None,
        "expected_type": expected[:1].decode("latin1") if expected else None,
        "byte_difference_count": len(expected),
        "actual_hash": None,
        "expected_hash": _sha256(expected),
        "fields": [],
        "intervals": [{"field": "REGISTRO_AUSENTE", "start": 1, "end": len(expected)}],
    }


def _extra_record(actual: bytes, actual_number: int) -> dict[str, Any]:
    return {
        "status": "extra",
        "record": actual_number,
        "actual_record": actual_number,
        "expected_record": None,
        "actual_type": actual[:1].decode("latin1") if actual else None,
        "expected_type": None,
        "byte_difference_count": len(actual),
        "actual_hash": _sha256(actual),
        "expected_hash": None,
        "fields": [],
        "intervals": [{"field": "REGISTRO_EXTRA", "start": 1, "end": len(actual)}],
    }


def _compare_record_block(
    actual_records: list[bytes],
    expected_records: list[bytes],
    *,
    actual_offset: int,
    expected_offset: int,
) -> list[dict]:
    differences: list[dict[str, Any]] = []
    matcher = SequenceMatcher(None, expected_records, actual_records, autojunk=False)
    for operation, expected_start, expected_end, actual_start, actual_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        expected_block = expected_records[expected_start:expected_end]
        actual_block = actual_records[actual_start:actual_end]
        paired = min(len(expected_block), len(actual_block))
        for offset in range(paired):
            differences.append(
                _changed_record(
                    actual_block[offset],
                    expected_block[offset],
                    actual_offset + actual_start + offset + 1,
                    expected_offset + expected_start + offset + 1,
                )
            )
        for offset, expected in enumerate(expected_block[paired:], start=paired):
            differences.append(
                _missing_record(expected, expected_offset + expected_start + offset + 1)
            )
        for offset, actual in enumerate(actual_block[paired:], start=paired):
            differences.append(_extra_record(actual, actual_offset + actual_start + offset + 1))
    return differences


def _compare_records(actual_records: list[bytes], expected_records: list[bytes]) -> list[dict]:
    # Keep physical wrappers aligned even when their count/sequence fields differ because
    # a detail was inserted or removed. This prevents the changed trailer from being
    # misclassified as the extra detail.
    has_wrappers = (
        len(actual_records) >= 2
        and len(expected_records) >= 2
        and actual_records[0][:1] == expected_records[0][:1] == b"0"
        and actual_records[-1][:1] == expected_records[-1][:1] == b"9"
    )
    if not has_wrappers:
        return _compare_record_block(
            actual_records, expected_records, actual_offset=0, expected_offset=0
        )

    differences: list[dict[str, Any]] = []
    if actual_records[0] != expected_records[0]:
        differences.append(_changed_record(actual_records[0], expected_records[0], 1, 1))
    differences.extend(
        _compare_record_block(
            actual_records[1:-1],
            expected_records[1:-1],
            actual_offset=1,
            expected_offset=1,
        )
    )
    if actual_records[-1] != expected_records[-1]:
        differences.append(
            _changed_record(
                actual_records[-1],
                expected_records[-1],
                len(actual_records),
                len(expected_records),
            )
        )
    return differences


def _parse_diagnostic(payload: bytes) -> tuple[bool, list[str]]:
    try:
        parse_cnab(payload)
    except CNABParseError as exc:
        return False, list(exc.errors)
    except (TypeError, ValueError) as exc:
        return False, [str(exc)]
    return True, []


def compare_cnab(actual: bytes, expected: bytes) -> dict[str, Any]:
    """Compare two payloads and return a JSON-serializable, privacy-safe report.

    Malformed input is represented under ``validation`` and still receives a literal
    byte/record comparison. ``record_differences`` uses one-based coordinates. Each
    field difference contains its full field bounds, exact changed intervals and hashes
    of the old/new field bytes; missing/extra records contain only whole-record hashes.
    """

    if not isinstance(actual, bytes) or not isinstance(expected, bytes):
        raise TypeError("actual e expected devem ser bytes")
    actual_valid, actual_errors = _parse_diagnostic(actual)
    expected_valid, expected_errors = _parse_diagnostic(expected)
    actual_records = _raw_records(actual)
    expected_records = _raw_records(expected)
    actual_hashes = [_sha256(record) for record in actual_records]
    expected_hashes = [_sha256(record) for record in expected_records]
    order_difference = (
        len(actual_hashes) == len(expected_hashes)
        and Counter(actual_hashes) == Counter(expected_hashes)
        and actual_hashes != expected_hashes
    )
    return {
        "equal": actual == expected and actual_valid and expected_valid,
        "byte_difference_count": _byte_difference_count(actual, expected),
        "lengths": {"actual": len(actual), "expected": len(expected)},
        "hashes": {"actual": _sha256(actual), "expected": _sha256(expected)},
        "record_counts": {"actual": len(actual_records), "expected": len(expected_records)},
        "order_difference": order_difference,
        "alignment_scope": "byte_content_and_position_not_credit_identity",
        "validation": {
            "actual": {"valid": actual_valid, "errors": actual_errors},
            "expected": {"valid": expected_valid, "errors": expected_errors},
        },
        "record_differences": _compare_records(actual_records, expected_records),
    }


def compare_linked_titles(actual: bytes, expected: bytes, record_pairs) -> dict[str, Any]:
    """Compare approved provenance pairs, with 1-based physical record coordinates.

    Unpaired details stay inconclusive. This never infers identity from amounts,
    names or order, and complements rather than masks the literal file comparison.
    """
    left, right = parse_cnab(actual)["records"], parse_cnab(expected)["records"]
    pairs = sorted(record_pairs)
    if len({a for a, _ in pairs}) != len(pairs) or len({b for _, b in pairs}) != len(pairs):
        raise ValueError("Vínculos de títulos não são um para um")
    if any(not (2 <= a < len(left) and 2 <= b < len(right)) for a, b in pairs):
        raise ValueError("Vínculo fora dos detalhes CNAB")
    return {
        "linked_titles": len(pairs),
        "order_difference": [b for _, b in pairs] != sorted(b for _, b in pairs),
        "actual_unlinked": sorted(set(range(2, len(left))) - {a for a, _ in pairs}),
        "expected_unlinked": sorted(set(range(2, len(right))) - {b for _, b in pairs}),
        "record_differences": [
            _changed_record(left[a - 1], right[b - 1], a, b)
            for a, b in pairs
            if left[a - 1] != right[b - 1]
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("actual", type=argparse.FileType("rb"))
    parser.add_argument("expected", type=argparse.FileType("rb"))
    args = parser.parse_args()
    report = compare_cnab(args.actual.read(), args.expected.read())
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
