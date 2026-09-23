"""Lectura de facturas con Vision LLM local vía Ollama (sin API cloud)."""

from __future__ import annotations

import base64
import io
import json
import os
import re
from pathlib import Path
from typing import Any

try:
    import requests
except Exception:  # pragma: no cover
    requests = None

DEFAULT_OLLAMA_BASE_URL = "http://127.0.0.1:11434"
DEFAULT_OLLAMA_VISION_MODEL = "qwen2.5vl:3b"
VISION_TIMEOUT_SEC = 180


def _load_dotenv_if_present() -> None:
    """Carga .env del proyecto sin dependencia extra (solo claves OLLAMA_* / TESSERACT_*)."""
    root = Path(__file__).resolve().parents[2]
    env_path = root / ".env"
    if not env_path.exists():
        return
    try:
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val
    except Exception:
        return


_load_dotenv_if_present()

INVOICE_VISION_PROMPT = """Eres un extractor de líneas de factura / listado de compra en español (Colombia).
Analiza la imagen y devuelve SOLO un JSON válido (sin markdown, sin explicaciones) con esta forma:

{
  "items": [
    {
      "code": "referencia o código si aparece",
      "description": "nombre del producto",
            "quantity": 0,
            "bonus_quantity": 0,
      "unit_price": 0,
      "total": 0,
      "discount_pct": null,
      "is_bonus": false
    }
  ],
  "raw_text": "texto útil opcional de la factura"
}

Reglas estrictas:
- Incluye SOLO líneas que representen productos o ítems comprados, nunca encabezados, títulos, subtotales, totales, impuestos, fechas, números de factura, cliente, proveedor, NIT, dirección, teléfono, pago, resumen ni datos de control.
- Si una línea es 'FACTURA', 'TOTAL', 'SUBTOTAL', 'IVA', 'RETENCIÓN', 'CLIENTE', 'NIT', 'FECHA', 'DIRECCIÓN', 'PAGOS', 'RESUMEN', 'ENTREGADO', 'BODEGA' o similar, NO la incluyas.
- Un producto real suele tener nombre, cantidad y valor unitario o total. Si una línea no parece un producto, omítela aunque tenga números.
- quantity, unit_price y total son números (usa punto decimal).
- Respeta los encabezados de la tabla: usa la columna "Cantidad" como quantity. Ignora columnas como inventario, existencia, disponible, descuento o código interno si no corresponden a cantidad.
- Comprueba cada línea con total = quantity pagada * unit_price. Si otra columna numérica parece una existencia, no la uses como quantity.
- quantity es la cantidad pagada; bonus_quantity es la cantidad bonificada/gratis de esa misma línea.
- Si hay 10 unidades pagadas y 3 bonificadas, devuelve quantity=10, bonus_quantity=3 y total=el valor pagado por las 10.
- is_bonus=true si la línea es obsequio/bonificación/gratis o lleva *.
- Si un campo no se ve, usa null o "" según corresponda.
- No inventes productos que no aparezcan en la imagen.
- Si hay duda entre un campo de producto y un campo genérico, descarta el genérico.
"""


def build_vision_prompt(manual_rules: dict[str, Any] | None = None) -> str:
    """Append manual extraction instructions to the image-based extraction prompt."""
    rules = manual_rules or {}
    code_cols = [
        item.strip()
        for item in str(rules.get("code_columns") or "").split(",")
        if item.strip()
    ] or ["Referencia", "Codigo", "Código", "GTIN", "EAN"]
    qty_cols = [
        item.strip()
        for item in str(rules.get("quantity_columns") or "").split(",")
        if item.strip()
    ] or ["Cantidad", "Cant.", "Cant"]
    price_cols = [
        item.strip()
        for item in str(rules.get("unit_price_columns") or "").split(",")
        if item.strip()
    ] or ["Precio Unitario", "Precio ud.", "Precio", "Valor Unitario"]
    discount_cols = [
        item.strip()
        for item in str(rules.get("discount_columns") or "").split(",")
        if item.strip()
    ] or ["Descuento", "Desc.", "Dto.", "% Descuento", "% Desc"]
    bonus_markers = [
        item.strip()
        for item in str(rules.get("bonus_markers") or "*, bonificación, obsequio, gratis, regalo, bono").split(",")
        if item.strip()
    ]
    return (
        f"\nReglas manuales del usuario:\n"
        f"- Busca código en columnas: {', '.join(code_cols)}.\n"
        f"- Busca cantidad en columnas: {', '.join(qty_cols)}.\n"
        f"- Busca precio unitario en columnas: {', '.join(price_cols)}.\n"
        f"- Busca descuento en columnas: {', '.join(discount_cols)}.\n"
        f"- Si una línea incluye alguno de estos indicadores de bonificación: {', '.join(bonus_markers)}, "
        "marca is_bonus=true y usa bonus_quantity para la cantidad gratis.\n"
        "- Si la línea lleva descuento o un asterisco asociado a la bonificación, corrige el costo "
        "considerando la cantidad total recibida (pagada + bonificada) y el descuento real."
    )


def ollama_base_url() -> str:
    return (os.environ.get("OLLAMA_BASE_URL") or DEFAULT_OLLAMA_BASE_URL).rstrip("/")


def ollama_vision_model() -> str:
    return (os.environ.get("OLLAMA_VISION_MODEL") or DEFAULT_OLLAMA_VISION_MODEL).strip()


def ollama_available(
    timeout: float = 2.0, model: str | None = None
) -> tuple[bool, str]:
    """Comprueba si el daemon Ollama responde en localhost."""
    if requests is None:
        return False, "Falta el paquete requests (pip install requests)."
    try:
        r = requests.get(f"{ollama_base_url()}/api/tags", timeout=timeout)
        if r.status_code != 200:
            return False, f"Ollama respondió HTTP {r.status_code}."
        models = [m.get("name", "") for m in (r.json().get("models") or [])]
        model_name = (model or ollama_vision_model()).strip()
        installed = any(
            name == model_name or name.startswith(f"{model_name}:") for name in models
        )
        if not installed:
            return (
                False,
                f"Ollama está activo pero falta el modelo '{model_name}'. "
                f"Ejecute: ollama pull {model_name}",
            )
        return True, f"Ollama listo ({model_name})"
    except Exception as exc:
        return (
            False,
            "Ollama no está disponible en "
            f"{ollama_base_url()}. Instale https://ollama.com y ejecute "
            f"`ollama pull {ollama_vision_model()}`. Detalle: {exc}",
        )


def _pil_to_png_b64(img) -> str:
    buf = io.BytesIO()
    if getattr(img, "mode", None) not in ("RGB", "L"):
        img = img.convert("RGB")
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _extract_json_object(text: str) -> dict[str, Any]:
    if not text:
        raise ValueError("Respuesta vacía del Vision LLM.")
    cleaned = text.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", cleaned, flags=re.I)
    if fence:
        cleaned = fence.group(1).strip()
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict):
            return data
        if isinstance(data, list):
            return {"items": data}
    except json.JSONDecodeError:
        pass
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start >= 0 and end > start:
        data = json.loads(cleaned[start : end + 1])
        if isinstance(data, dict):
            return data
    raise ValueError("No se pudo parsear JSON de la respuesta Vision.")


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "si", "sí", "yes", "y", "*"}


def _as_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("$", "").replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _reconcile_quantity(quantity: Any, unit: Any, total: Any) -> Any:
    """Correct a quantity when another table column was read as quantity."""
    qty = _as_number(quantity)
    unit_value = _as_number(unit)
    total_value = _as_number(total)
    if not qty or not unit_value or not total_value or unit_value <= 0:
        return quantity
    inferred = total_value / unit_value
    if inferred <= 0 or abs(inferred - round(inferred)) > 1e-6:
        return quantity
    if abs(qty - inferred) > 1e-6:
        return int(round(inferred))
    return quantity


def _is_header_or_non_product_description(desc: str) -> bool:
    if not desc:
        return True
    text = re.sub(r"\s+", " ", desc.strip())
    if len(text) <= 2:
        return True
    normalized = text.lower()

    metadata_prefixes = (
        "factura", "remision", "remisión", "subtotal", "total", "valor total",
        "gran total", "resumen", "iva", "impuesto", "retencion", "retención",
        "cliente", "proveedor", "nit", "direccion", "tel", "telefono", "fecha",
        "hora", "pedido", "vendedor", "bodega", "caja", "forma de pago",
        "efectivo", "transferencia", "tarjeta", "pagos", "condiciones",
        "observacion", "observación", "encabezado", "nota"
    )

    if re.fullmatch(r"[\d\s\$\.\-,%/]+", text):
        return True

    if re.match(r"^(cliente|proveedor|nit|fecha|hora|direccion|tel|telefono|pedido|vendedor)\s*[:\-]", normalized):
        return True

    if any(normalized.startswith(prefix) for prefix in metadata_prefixes):
        return True

    if any(normalized.startswith(prefix + " ") for prefix in ("total", "subtotal", "iva", "impuesto", "cliente", "cte", "proveedor", "nit", "bodega", "resumen", "pagos")):
        return True

    if re.search(r"\b(?:subtotal|total|iva|impuesto|retencion|retención|cliente|proveedor|nit|direccion|telefono|fecha|hora|pedido|vendedor|bodega|resumen|pagos)\b", normalized):
        if len(normalized.split()) <= 6:
            return True

    return False


def normalize_vision_items(payload: dict[str, Any], source_name: str = "") -> list[dict]:
    """Convierte el JSON del modelo al esquema de filas de app_effi."""
    items = payload.get("items")
    if items is None and isinstance(payload.get("lineas"), list):
        items = payload["lineas"]
    if not isinstance(items, list):
        items = []

    rows: list[dict] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        desc = str(
            item.get("description")
            or item.get("descripcion")
            or item.get("nombre")
            or ""
        ).strip()
        if not desc or _is_header_or_non_product_description(desc):
            continue
        code = str(item.get("code") or item.get("codigo") or item.get("referencia") or "").strip()
        qty = item.get("quantity", item.get("cantidad"))
        bonus_qty = item.get(
            "bonus_quantity",
            item.get("cantidad_bonificada", item.get("cantidad_bonus", 0)),
        )
        unit = item.get("unit_price", item.get("precio_unitario", item.get("precio")))
        total = item.get("total", item.get("valor", item.get("importe")))
        qty = _reconcile_quantity(qty, unit, total)
        disc = item.get("discount_pct", item.get("descuento_pct", item.get("descuento")))
        is_bonus = _as_bool(item.get("is_bonus", item.get("bonificacion", item.get("obsequio"))))
        rows.append(
            {
                "archivo_origen": source_name,
                "code": code,
                "description": desc,
                "presentation": str(item.get("presentation") or item.get("presentacion") or ""),
                "quantity": qty,
                "bonus_quantity": bonus_qty,
                "unit_price_invoice": unit,
                "total_invoice": total,
                "discount_pct": disc,
                "discount_value": item.get("discount_value") or "",
                "is_bonus": is_bonus,
                "source_line": f"[vision] {code} {desc} qty={qty} unit={unit} total={total}".strip(),
            }
        )
    return rows


def extract_invoice_from_image_b64(
    image_b64: str,
    *,
    model: str | None = None,
    timeout: int = VISION_TIMEOUT_SEC,
    manual_rules: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if requests is None:
        raise RuntimeError("Instale requests: pip install requests")
    model_name = model or ollama_vision_model()
    url = f"{ollama_base_url()}/api/chat"
    body = {
        "model": model_name,
        "stream": False,
        "format": "json",
        "messages": [
            {
                "role": "user",
                "content": INVOICE_VISION_PROMPT + build_vision_prompt(manual_rules),
                "images": [image_b64],
            }
        ],
        "options": {"temperature": 0},
    }
    try:
        r = requests.post(url, json=body, timeout=timeout)
    except requests.Timeout as exc:
        raise TimeoutError(
            f"Vision LLM agotó el tiempo ({timeout}s) con modelo {model_name}."
        ) from exc
    except requests.RequestException as exc:
        raise RuntimeError(f"Error llamando a Ollama: {exc}") from exc

    if r.status_code != 200:
        detail = r.text[:500]
        raise RuntimeError(
            f"Ollama HTTP {r.status_code} (modelo={model_name}). {detail}"
        )

    data = r.json()
    content = ""
    msg = data.get("message") or {}
    if isinstance(msg, dict):
        content = msg.get("content") or ""
    if not content:
        content = data.get("response") or ""
    return _extract_json_object(content)


def extract_invoice_rows_from_pil(
    img,
    *,
    source_name: str = "",
    model: str | None = None,
    manual_rules: dict[str, Any] | None = None,
    timeout: int = VISION_TIMEOUT_SEC,
) -> tuple[list[dict], str]:
    """Devuelve (filas_normalizadas, texto_auditoria)."""
    payload = extract_invoice_from_image_b64(
        _pil_to_png_b64(img),
        model=model,
        manual_rules=manual_rules,
        timeout=timeout,
    )
    rows = normalize_vision_items(payload, source_name=source_name)
    raw_text = str(payload.get("raw_text") or "").strip()
    if not raw_text:
        raw_text = json.dumps(payload, ensure_ascii=False, indent=2)
    return rows, raw_text


def extract_invoice_rows_from_images(
    images: list,
    *,
    source_name: str = "",
    model: str | None = None,
    manual_rules: dict[str, Any] | None = None,
    timeout: int = VISION_TIMEOUT_SEC,
) -> tuple[list[dict], str]:
    all_rows: list[dict] = []
    texts: list[str] = []
    for idx, img in enumerate(images, 1):
        rows, text = extract_invoice_rows_from_pil(
            img,
            source_name=source_name,
            model=model,
            manual_rules=manual_rules,
            timeout=timeout,
        )
        all_rows.extend(rows)
        texts.append(f"--- PÁGINA {idx} (Vision) ---\n{text}")
    return all_rows, "\n".join(texts)
