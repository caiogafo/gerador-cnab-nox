"""Independent literal positional oracle from Specs §6; no application imports."""


def expected_payload():
    header = bytearray(b" " * 444)
    detail = bytearray(b" " * 444)
    trailer = bytearray(b" " * 444)

    def put(record, start, end, content):
        assert len(content) == end - start + 1
        record[start - 1 : end] = content

    for start, end, value in [
        (1, 1, b"0"),
        (2, 2, b"1"),
        (3, 9, b"REMESSA"),
        (10, 11, b"01"),
        (12, 26, b"COBRANCA       "),
        (27, 46, b"00000000000000000125"),
        (47, 76, b"0" * 30),
        (77, 79, b"001"),
        (80, 94, b"0" * 15),
        (95, 100, b"030926"),
        (109, 110, b"MX"),
        (111, 117, b"0000001"),
        (439, 444, b"000001"),
    ]:
        put(header, start, end, value)
    for start, end, value in [
        (1, 1, b"1"),
        (21, 22, b"02"),
        (23, 37, b"0" * 15),
        (38, 62, b" " * 22 + b"123"),
        (63, 65, b"001"),
        (66, 70, b"0" * 5),
        (71, 81, b"0" * 11),
        (82, 82, b"1"),
        (83, 92, b"0" * 10),
        (93, 93, b"1"),
        (94, 94, b"N"),
        (95, 100, b"030926"),
        (106, 106, b"1"),
        (109, 110, b"01"),
        (111, 120, b" " * 7 + b"123"),
        (121, 126, b"310130"),
        (127, 139, b"0000000015000"),
        (140, 142, b"0" * 3),
        (143, 147, b"0" * 5),
        (148, 149, b"24"),
        (151, 156, b"010926"),
        (157, 159, b"0" * 3),
        (160, 161, b"01"),
        (162, 192, b"0" * 31),
        (193, 205, b"0000000010000"),
        (206, 218, b"0" * 13),
        (219, 220, b"02"),
        (221, 234, b"11222333000181"),
        (235, 274, b"FALENCIA ALFA SOCIEDADE ANONIMA".ljust(40)),
        (327, 334, b"0" * 8),
        (335, 380, b"ACME LIMITADA".ljust(46)),
        (381, 394, b"   52998224725"),
        (395, 438, b"0" * 44),
        (439, 444, b"000002"),
    ]:
        put(detail, start, end, value)
    put(trailer, 1, 1, b"9")
    put(trailer, 439, 444, b"000003")
    return b"\r\n".join((header, detail, trailer)) + b"\r\n"
