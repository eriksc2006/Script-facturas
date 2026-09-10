from effi_processor.quality_rules import (
    EFFI_HEADERS,
    calculate_discount_value,
    calculate_prorated_unit_cost,
    convert_quantity,
    has_exact_effI_headers,
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
    # 10 paid + 2 free units, total paid = 100
    assert calculate_prorated_unit_cost(100, 10, 2) == 100 / 12


def test_full_bonus_discount():
    assert calculate_discount_value(50, 100) == 50


def test_zero_discount():
    assert calculate_discount_value(50, 0) == 0


def test_unit_conversion():
    # One 4L presentation equals 4 / 3.785411784 galones.
    result = convert_quantity(1, 4, 3.785411784)
    assert abs(result - (4 / 3.785411784)) < 1e-9
