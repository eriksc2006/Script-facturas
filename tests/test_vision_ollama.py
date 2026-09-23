"""Tests for local Ollama vision helpers (no daemon required)."""

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

sys.path.insert(0, str(ROOT))

from app_effi import match_catalog_item  # noqa: E402
from effi_processor.vision_ollama import (  # noqa: E402
    _extract_json_object,
    normalize_vision_items,
)


def test_extract_json_from_fenced_response():
    raw = """```json
{"items": [{"description": "SHAMPOO", "quantity": 2, "unit_price": 1000, "total": 2000}]}
```"""
    data = _extract_json_object(raw)
    assert data["items"][0]["description"] == "SHAMPOO"


def test_normalize_vision_items_maps_schema():
    payload = {
        "items": [
            {
                "code": "TA91",
                "description": "ACTIVADOR THYMS",
                "quantity": 10,
                "bonus_quantity": 3,
                "unit_price": 2500,
                "total": 25000,
                "is_bonus": False,
            },
            {"description": "", "quantity": 1},
        ]
    }
    rows = normalize_vision_items(payload, source_name="factura.png")
    assert len(rows) == 1
    assert rows[0]["code"] == "TA91"
    assert rows[0]["archivo_origen"] == "factura.png"
    assert rows[0]["unit_price_invoice"] == 2500
    assert rows[0]["bonus_quantity"] == 3
    assert rows[0]["source_line"].startswith("[vision]")


def test_normalize_accepts_spanish_keys():
    payload = {
        "items": [
            {
                "codigo": "X1",
                "descripcion": "PRODUCTO",
                "cantidad": 3,
                "precio_unitario": 100,
                "valor": 300,
                "bonificacion": True,
            }
        ]
    }
    rows = normalize_vision_items(payload)
    assert rows[0]["is_bonus"] is True
    assert rows[0]["quantity"] == 3


def test_normalize_corrects_quantity_using_invoice_total():
    # Caso de la factura fotografiada: una columna numérica adicional puede
    # leerse como 10, pero 56.800 / 28.400 demuestra que la cantidad es 2.
    payload = {
        "items": [
            {
                "code": "SILMANTSH",
                "description": "SILKY MANTENIMIENTO SHAMPOO COLOR CARE",
                "presentation": "250ml",
                "quantity": 10,
                "unit_price": 28400,
                "total": 56800,
            }
        ]
    }
    rows = normalize_vision_items(payload)
    assert rows[0]["quantity"] == 2


def test_normalize_keeps_paid_quantity_for_bonus_rule():
    payload = {
        "items": [
            {
                "description": "PRODUCTO BONIFICADO",
                "quantity": 10,
                "bonus_quantity": 3,
                "unit_price": 100,
                "total": 1000,
            }
        ]
    }
    rows = normalize_vision_items(payload)
    assert rows[0]["quantity"] == 10
    assert rows[0]["bonus_quantity"] == 3


def test_normalize_filters_headers_and_totals_not_products():
    payload = {
        "items": [
            {"description": "FACTURA ELECTRÓNICA", "quantity": 1, "unit_price": 1, "total": 1},
            {"description": "TOTAL", "quantity": 1, "unit_price": 50000, "total": 50000},
            {"description": "SUBTOTAL", "quantity": 1, "unit_price": 40000, "total": 40000},
            {"description": "IVA 19%", "quantity": 1, "unit_price": 7600, "total": 7600},
            {"description": "ACEITE VEGETAL 1L", "quantity": 2, "unit_price": 8500, "total": 17000},
            {"description": "CLIENTE: JUAN PEREZ", "quantity": 1, "unit_price": 0, "total": 0},
        ]
    }
    rows = normalize_vision_items(payload, source_name="factura.png")
    assert len(rows) == 1
    assert rows[0]["description"] == "ACEITE VEGETAL 1L"
    assert rows[0]["quantity"] == 2


def test_match_catalog_item_finds_similar_inventory_product():
    catalog = pd.DataFrame([
        {
            "Código de barras GTIN": "7701234567890",
            "Descripción": "ACEITE VEGETAL 1 LITRO",
            "Presentación": "BOTELLA 1 L",
        },
        {
            "Código de barras GTIN": "7709876543210",
            "Descripción": "ARROZ 1 KG",
            "Presentación": "BOLSA 1 KG",
        },
    ])
    cols = {
        "gtin": "Código de barras GTIN",
        "description": "Descripción",
        "presentation": "Presentación",
    }
    match_row, score, method = match_catalog_item(
        "ACEITE VEGETAL 1 LT",
        "BOTELLA",
        catalog,
        cols,
    )
    assert match_row is not None
    assert score >= 0.6
    assert "ACEITE" in str(match_row["Descripción"]).upper()
    assert method in {"Fuzzy", "token"}
