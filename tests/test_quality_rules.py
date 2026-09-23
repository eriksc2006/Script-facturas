from effi_processor.quality_rules import (
    EFFI_HEADERS,
    calculate_discount_value,
    calculate_prorated_unit_cost,
    convert_quantity,
    has_exact_effI_headers,
    is_bonus_marker,
    normalize_manual_list,
    validate_main_sheet_rows,
)


def test_effI_headers_are_exact():
    assert has_exact_effI_headers(EFFI_HEADERS)


def test_effI_headers_reject_modified_accent_or_order():
    modified = EFFI_HEADERS.copy()
    modified[0] = "Articulo (ID EFFI | Código de barras GTIN | Serie)"
    assert not has_exact_effI_headers(modified)


def test_main_sheet_requires_nine_columns():
    assert validate_main_sheet_rows([[None] * 9])
    assert not validate_main_sheet_rows([[None] * 8])


def test_bonus_prorating():
    # 10 paid + 3 free units, total paid = 100
    assert calculate_prorated_unit_cost(100, 10, 3) == 100 / 13


def test_full_bonus_discount():
    assert calculate_discount_value(50, 100) == 50


def test_zero_discount():
    assert calculate_discount_value(50, 0) == 0


def test_unit_conversion():
    # One 4L presentation equals 4 / 3.785411784 galones.
    result = convert_quantity(1, 4, 3.785411784)
    assert abs(result - (4 / 3.785411784)) < 1e-9


def test_manual_list_normalizes_values():
    assert normalize_manual_list("Cantidad; Cant., Cant") == [
        "Cantidad",
        "Cant.",
        "Cant",
    ]


def test_bonus_marker_uses_manual_rules():
    assert is_bonus_marker("Línea con * regalo", ["*", "regalo"]) is True
    assert is_bonus_marker("Línea normal", ["*", "regalo"]) is False


def test_bonus_cost_is_prorated_after_discount():
    unit_cost = calculate_prorated_unit_cost(1000, 10, 2, discount_value=100)
    assert abs(unit_cost - (900 / 12)) < 1e-9
