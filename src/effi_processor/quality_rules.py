"""Pure validation functions used by the Effi processor and its tests."""

EFFI_HEADERS = [
    "Artículo (ID EFFI | Código de barras GTIN | Serie)",
    "ID Tipo de Egreso",
    "Lote",
    "Serie",
    "Observación",
    "Cantidad *",
    "Precio ud. *",
    "Valor descuento total. *",
    "Código Effi Impuesto",
]


def has_exact_effI_headers(headers: list[str]) -> bool:
    """Return True only when the import headers match the required order exactly."""
    return headers == EFFI_HEADERS


def validate_main_sheet_rows(rows: list[list[object]]) -> bool:
    """Validate that each import row has exactly nine columns."""
    return all(len(row) == 9 for row in rows)


def calculate_prorated_unit_cost(total_paid: float, paid_qty: float, bonus_qty: float) -> float:
    """Spread the amount paid over all physical units received."""
    total_qty = paid_qty + bonus_qty
    if total_qty <= 0:
        raise ValueError("Total physical quantity must be greater than zero.")
    return total_paid / total_qty


def calculate_discount_value(base_value: float, discount_pct: float) -> float:
    """Calculate an item's discount without applying a global discount."""
    if not 0 <= discount_pct <= 100:
        raise ValueError("Discount percentage must be between 0 and 100.")
    return base_value * discount_pct / 100.0


def convert_quantity(quantity: float, document_factor: float, catalog_factor: float) -> float:
    """Convert a document quantity into the catalog presentation."""
    if quantity < 0:
        raise ValueError("Quantity cannot be negative.")
    if document_factor <= 0 or catalog_factor <= 0:
        raise ValueError("Presentation factors must be positive.")
    return quantity * document_factor / catalog_factor
