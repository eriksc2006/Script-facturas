"""Pure validation functions used by the Effi processor and its tests."""

import re

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


def normalize_manual_list(value: str | None) -> list[str]:
    """Split custom values like 'Cantidad; Cant.; Cant' into a clean list."""
    if value is None:
        return []
    items = re.split(r"[;,|\n]+", str(value))
    cleaned: list[str] = []
    for item in items:
        candidate = item.strip().strip('"').strip("'")
        if candidate and candidate not in cleaned:
            cleaned.append(candidate)
    return cleaned


def is_bonus_marker(text: str | None, markers: list[str] | tuple[str, ...] | None = None) -> bool:
    """Detect bonus indicators including the '*' symbol and custom text markers."""
    raw = str(text or "")
    explicit = list(markers or ["*", "bonificacion", "obsequio", "gratis", "regalo", "bono"])
    if not explicit:
        explicit = ["*", "bonificacion", "obsequio", "gratis", "regalo", "bono"]
    needle = raw.strip().lower()
    for marker in explicit:
        candidate = str(marker).strip().lower()
        if not candidate:
            continue
        if candidate == "*" and "*" in raw:
            return True
        if candidate in needle:
            return True
    return False


def calculate_prorated_unit_cost(
    total_paid: float,
    paid_qty: float,
    bonus_qty: float,
    discount_pct: float = 0.0,
    discount_value: float = 0.0,
) -> float:
    """Spread the net amount paid over all physical units received."""
    discount_amount = float(discount_value or 0.0)
    if discount_amount <= 0 and float(discount_pct or 0.0) > 0:
        discount_amount = float(total_paid) * (float(discount_pct) / 100.0)
    net_paid = max(float(total_paid) - discount_amount, 0.0)
    total_qty = float(paid_qty) + float(bonus_qty)
    if total_qty <= 0:
        raise ValueError("Total physical quantity must be greater than zero.")
    return net_paid / total_qty


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
