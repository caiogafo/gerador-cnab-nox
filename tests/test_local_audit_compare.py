from __future__ import annotations

import json
from datetime import date
from decimal import Decimal

import pytest

from scripts.local_audit_compare import (
    FIELD_RANGES,
    CNABParseError,
    compare_cnab,
    compare_linked_titles,
    parse_cnab,
)
from tests.golden_expected import expected_payload


def _records(payload: bytes) -> list[bytearray]:
    return [bytearray(record) for record in payload[:-2].split(b"\r\n")]


def _payload(records: list[bytearray]) -> bytes:
    return b"\r\n".join(bytes(record) for record in records) + b"\r\n"


def _put(record: bytearray, start: int, end: int, value: bytes) -> None:
    assert len(value) == end - start + 1
    record[start - 1 : end] = value


def _two_detail_payload() -> bytes:
    records = _records(expected_payload())
    second = bytearray(records[1])
    _put(second, 38, 62, b" " * 22 + b"124")
    _put(second, 111, 120, b" " * 7 + b"124")
    _put(second, 335, 374, b"OUTRA EMPRESA".ljust(40))
    _put(second, 439, 444, b"000003")
    _put(records[2], 439, 444, b"000004")
    return _payload([records[0], records[1], second, records[2]])


def _field_entry(report: dict, field: str) -> dict:
    fields = [
        item
        for record in report["record_differences"]
        for item in record["fields"]
        if item["field"] == field
    ]
    assert len(fields) == 1
    return fields[0]


def test_parse_cnab_returns_raw_records_and_typed_21_fields() -> None:
    parsed = parse_cnab(expected_payload())

    assert len(parsed["records"]) == 3
    assert parsed["records"][0][:1] == b"0"
    assert parsed["records"][1][:1] == b"1"
    assert parsed["records"][2][:1] == b"9"
    assert len(parsed["details"]) == 1
    assert tuple(parsed["details"][0]) == tuple(FIELD_RANGES)
    assert parsed["liquidation_date"] == date(2026, 9, 3)
    assert parsed["first_sequence"] == 123

    detail = parsed["details"][0]
    assert detail["DOC_CEDENTE"] == "52998224725"
    assert detail["NOME_CEDENTE"] == "ACME LIMITADA"
    assert detail["DT_VENCIMENTO"] == date(2030, 1, 31)
    assert detail["VL_NOMINAL"] == Decimal("150")
    assert detail["VL_PRESENTE"] == Decimal("100")
    assert detail["TIPO_PESSOA_CEDENTE"] == 1
    assert detail["TIPO_PESSOA_SACADO"] == 2
    assert detail["TAXA_INDEXADOR"] is None


def test_equal_comparison_is_json_serializable() -> None:
    payload = expected_payload()
    report = compare_cnab(payload, payload)

    assert report["equal"] is True
    assert report["byte_difference_count"] == 0
    assert report["record_differences"] == []
    assert report["validation"]["actual"] == {"valid": True, "errors": []}
    json.dumps(report)


def test_name_difference_is_not_ignored_and_never_exposes_contents() -> None:
    expected = expected_payload()
    records = _records(expected)
    records[1][334] = ord("Z")
    report = compare_cnab(_payload(records), expected)

    name = _field_entry(report, "NOME_CEDENTE")
    assert name["start"] == 335 and name["end"] == 374
    assert name["byte_difference_count"] == 1
    assert name["intervals"] == [{"start": 335, "end": 335}]
    serialized = json.dumps(report, ensure_ascii=False)
    assert "ACME LIMITADA" not in serialized
    assert "ZCME LIMITADA" not in serialized


def test_one_cent_difference_is_attributed_to_present_value() -> None:
    expected = expected_payload()
    records = _records(expected)
    records[1][204] = ord("1")
    report = compare_cnab(_payload(records), expected)

    value = _field_entry(report, "VL_PRESENTE")
    assert value["intervals"] == [{"start": 205, "end": 205}]
    assert value["byte_difference_count"] == 1


def test_padding_difference_is_visible_as_padding() -> None:
    expected = expected_payload()
    records = _records(expected)
    records[1][1] = ord("X")
    report = compare_cnab(_payload(records), expected)

    padding = _field_entry(report, "PADDING_001")
    assert padding["intervals"] == [{"start": 2, "end": 2}]


def test_swapped_records_are_reported_as_order_difference() -> None:
    expected = _two_detail_payload()
    records = _records(expected)
    actual = _payload([records[0], records[2], records[1], records[3]])
    report = compare_cnab(actual, expected)

    assert report["equal"] is False
    assert report["order_difference"] is True
    statuses = [difference["status"] for difference in report["record_differences"]]
    assert statuses == ["extra", "missing"]


def test_reordered_and_renumbered_titles_use_provenance_without_false_name_differences():
    expected = _two_detail_payload()
    records = _records(expected)
    swapped = [records[0], records[2], records[1], records[3]]
    for physical, record in enumerate(swapped[1:-1], 2):
        sequence = str(121 + physical).encode()
        _put(record, 38, 62, sequence.rjust(25))
        _put(record, 111, 120, sequence.rjust(10))
        _put(record, 439, 444, str(physical).encode().zfill(6))
    actual = _payload(swapped)
    literal = compare_cnab(actual, expected)
    linked = compare_linked_titles(actual, expected, [(2, 3), (3, 2)])
    assert not literal["equal"] and literal["byte_difference_count"] > 0
    assert linked["order_difference"]
    assert linked["actual_unlinked"] == linked["expected_unlinked"] == []
    fields = {f["field"] for r in linked["record_differences"] for f in r["fields"]}
    assert fields == {"SEU_NUMERO", "NU_DOCUMENTO", "SEQUENCIA_FISICA"}


def test_missing_and_extra_records_have_one_based_coordinates_and_hashes() -> None:
    expected = expected_payload()
    records = _records(expected)
    missing_report = compare_cnab(_payload([records[0], records[2]]), expected)
    missing = next(
        item for item in missing_report["record_differences"] if item["status"] == "missing"
    )
    assert missing["expected_record"] == 2
    assert missing["actual_record"] is None
    assert missing["expected_hash"] and missing["actual_hash"] is None

    actual = _two_detail_payload()
    extra_report = compare_cnab(actual, expected)
    extra = next(item for item in extra_report["record_differences"] if item["status"] == "extra")
    assert extra["actual_record"] == 3
    assert extra["expected_record"] is None
    assert extra["actual_hash"] and extra["expected_hash"] is None


@pytest.mark.parametrize(
    ("mutate", "message"),
    [
        (lambda payload: b"\xef\xbb\xbf" + payload, "BOM não permitido"),
        (lambda payload: payload[:-2], "deve terminar com CRLF"),
        (lambda payload: payload.replace(b"\r\n", b"\n"), "fora do padrão CRLF"),
        (lambda payload: payload[:443] + payload[444:], "possui 443 bytes"),
    ],
)
def test_malformed_envelopes_raise_clear_parse_errors_and_compare_does_not_abort(
    mutate, message: str
) -> None:
    expected = expected_payload()
    actual = mutate(expected)

    with pytest.raises(CNABParseError, match=message):
        parse_cnab(actual)
    report = compare_cnab(actual, expected)
    assert report["equal"] is False
    assert report["validation"]["actual"]["valid"] is False
    assert any(message in error for error in report["validation"]["actual"]["errors"])
    assert report["byte_difference_count"] > 0


def test_invalid_cp1252_byte_is_diagnosed_without_field_contents() -> None:
    expected = expected_payload()
    records = _records(expected)
    records[1][334] = 0x81
    actual = _payload(records)

    report = compare_cnab(actual, expected)
    assert report["validation"]["actual"]["valid"] is False
    errors = report["validation"]["actual"]["errors"]
    assert errors == ["registro 2: byte incompatível com CP1252 na posição 335"]
    assert _field_entry(report, "NOME_CEDENTE")["intervals"] == [{"start": 335, "end": 335}]


def test_header_trailer_and_final_crlf_are_part_of_literal_comparison() -> None:
    expected = expected_payload()
    header_changed = bytearray(expected)
    header_changed[2] = ord("X")
    header_report = compare_cnab(bytes(header_changed), expected)
    assert _field_entry(header_report, "HEADER_LITERAL_REMESSA")["intervals"] == [
        {"start": 3, "end": 3}
    ]

    trailer_records = _records(expected)
    trailer_records[-1][1] = ord("X")
    trailer_report = compare_cnab(_payload(trailer_records), expected)
    assert _field_entry(trailer_report, "TRAILER_PADDING")["intervals"] == [{"start": 2, "end": 2}]

    newline_report = compare_cnab(expected[:-1], expected)
    assert newline_report["lengths"]["actual"] == len(expected) - 1
    assert newline_report["byte_difference_count"] == 1
