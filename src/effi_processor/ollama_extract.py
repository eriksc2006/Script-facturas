"""Local Ollama (free) extractor for unstructured invoice text."""

from __future__ import annotations

import json
import os
import re
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

DEFAULT_HOST = "http://127.0.0.1:11434"
DEFAULT_MODEL = "llama3.2"
DEFAULT_TIMEOUT = 120.0
MAX_TEXT_CHARS = 12000

SYSTEM_PROMPT = (
    "Eres un extractor de líneas de facturas (Colombia y Latinoamérica). "
    "Responde únicamente JSON válido con esta forma: "
    '{"tax_included": null, "items": ['
    '{"code": "", "description": "", "presentation": "", "quantity": 0, '
    '"unit_price_invoice": 0, "total_invoice": 0, "discount_pct": null, '
    '"discount_value": null, "is_bonus": false, "source_line": ""}]}'
    " Reglas: extrae solo productos o servicios del documento; no inventes datos; "
    "code solo si el texto trae referencia, SKU, GTIN o EAN; nunca inventes "
    "códigos Effi ni IDs internos; is_bonus es true si es bonificación, obsequio, "
    "gratis o línea con asterisco; usa punto decimal; omite encabezados, NIT, "
    "totales generales, IVA global, fletes y pie de página."
)


def default_host() -> str:
    return (os.environ.get("OLLAMA_HOST") or DEFAULT_HOST).strip().rstrip("/")


def default_model() -> str:
    return (os.environ.get("OLLAMA_MODEL") or DEFAULT_MODEL).strip()


def max_prompt_chars(model: str | None) -> int:
    """Keep prompts inside small local models such as moondream (~2k context)."""
    name = (model or default_model()).strip().lower()
    if "moondream" in name:
        return 2800
    return MAX_TEXT_CHARS


def validate_ollama_base(host: str) -> str:
    """Return a normalized http(s) base URL or raise ValueError."""
    base = (host or "").strip().rstrip("/")
    parsed = urlparse(base)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError(
            "OLLAMA_HOST debe ser una URL http(s), por ejemplo "
            "http://127.0.0.1:11434"
        )
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}".rstrip("/")


def _request_json(
    url: str,
    payload: dict[str, Any] | None = None,
    timeout: float = 5.0,
) -> Any:
    headers = {"Accept": "application/json"}
    data = None
    method = "GET"
    if payload is not None:
        method = "POST"
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = Request(url, data=data, headers=headers, method=method)  # noqa: S310
    with urlopen(request, timeout=timeout) as response:  # nosec B310
        raw = response.read().decode("utf-8", errors="replace")
    if not raw:
        return {}
    return json.loads(raw)


def check_ollama(
    host: str | None = None,
    timeout: float = 3.0,
) -> dict[str, Any]:
    """Probe a local Ollama daemon. Never requires a paid API key."""
    try:
        base = validate_ollama_base(host or default_host())
    except ValueError as exc:
        return {
            "ok": False,
            "host": host or default_host(),
            "models": [],
            "message": str(exc),
        }

    tags_url = urljoin(base + "/", "api/tags")
    try:
        body = _request_json(tags_url, timeout=timeout)
    except HTTPError as exc:
        return {
            "ok": False,
            "host": base,
            "models": [],
            "message": f"Ollama respondió HTTP {exc.code} en {base}.",
        }
    except URLError:
        return {
            "ok": False,
            "host": base,
            "models": [],
            "message": (
                "Ollama no está en ejecución. Instale la versión gratuita desde "
                "https://ollama.com, inicie la aplicación y ejecute "
                f"`ollama pull {default_model()}`."
            ),
        }
    except TimeoutError:
        return {
            "ok": False,
            "host": base,
            "models": [],
            "message": f"Tiempo de espera agotado al consultar Ollama en {base}.",
        }
    except (OSError, json.JSONDecodeError) as exc:
        return {
            "ok": False,
            "host": base,
            "models": [],
            "message": f"No se pudo leer el estado de Ollama: {exc}",
        }

    models = []
    for item in body.get("models") or []:
        name = item.get("name") if isinstance(item, dict) else None
        if name:
            models.append(str(name))
    if models:
        message = f"Ollama local listo ({base}). Modelos: {', '.join(models[:8])}"
        ready = True
    else:
        message = (
            f"Ollama está en {base} pero no hay modelos. "
            f"Ejecute `ollama pull {default_model()}`."
        )
        ready = False
    return {
        "ok": ready,
        "host": base,
        "models": models,
        "message": message,
    }


def parse_number_loose(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts[-1]) in {1, 2}:
            text = "".join(parts[:-1]).replace(".", "") + "." + parts[-1]
        else:
            text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def coerce_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "t", "si", "sí", "yes", "y"}


def extract_json_object(raw: str) -> dict[str, Any]:
    text = (raw or "").strip()
    if not text:
        return {}
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.S)
    if fenced:
        text = fenced.group(1)
    try:
        loaded = json.loads(text)
        return loaded if isinstance(loaded, dict) else {}
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start < 0 or end <= start:
            return {}
        try:
            loaded = json.loads(text[start : end + 1])
        except json.JSONDecodeError:
            return {}
        return loaded if isinstance(loaded, dict) else {}


def normalize_item(item: Any, source_name: str = "") -> dict[str, Any] | None:
    if not isinstance(item, dict):
        return None
    description = str(item.get("description") or item.get("descripcion") or "").strip()
    code = str(item.get("code") or item.get("codigo") or item.get("gtin") or "").strip()
    if not description and not code:
        return None
    qty = parse_number_loose(item.get("quantity", item.get("cantidad")))
    unit = parse_number_loose(
        item.get("unit_price_invoice", item.get("precio_unitario"))
    )
    total = parse_number_loose(item.get("total_invoice", item.get("total")))
    if qty is None:
        qty = 1.0
    presentation = str(item.get("presentation") or item.get("presentacion") or "")
    source_line = str(item.get("source_line") or description)
    return {
        "code": code,
        "description": description or code,
        "presentation": presentation.strip(),
        "quantity": qty,
        "unit_price_invoice": unit,
        "total_invoice": total,
        "discount_pct": parse_number_loose(item.get("discount_pct")),
        "discount_value": parse_number_loose(item.get("discount_value")),
        "is_bonus": coerce_bool(item.get("is_bonus")),
        "source_line": source_line,
        "archivo_origen": source_name,
        "extraction_source": "ollama",
    }


def items_from_payload(
    payload: dict[str, Any], source_name: str = ""
) -> list[dict[str, Any]]:
    raw_items = payload.get("items")
    if raw_items is None and isinstance(payload.get("lineas"), list):
        raw_items = payload["lineas"]
    if not isinstance(raw_items, list):
        return []
    rows: list[dict[str, Any]] = []
    for item in raw_items:
        normalized = normalize_item(item, source_name=source_name)
        if normalized:
            rows.append(normalized)
    return rows


def select_extracted_rows(
    regex_rows: list[dict[str, Any]],
    ollama_rows: list[dict[str, Any]],
    mode: str = "fallback",
) -> list[dict[str, Any]]:
    """Choose regex vs local LLM rows. `always` prefers Ollama when it found lines."""
    normalized_mode = (mode or "fallback").strip().lower()
    if normalized_mode == "always" and ollama_rows:
        return ollama_rows
    if regex_rows:
        return regex_rows
    return ollama_rows


def should_call_ollama(
    enabled: bool,
    mode: str,
    regex_rows: list[dict[str, Any]],
    text: str,
) -> bool:
    """Skip the local LLM when disabled, empty, or fallback already has regex lines."""
    if not enabled or not (text or "").strip():
        return False
    if (mode or "fallback").strip().lower() != "always" and regex_rows:
        return False
    return True


def extract_invoice_items(
    text: str,
    host: str | None = None,
    model: str | None = None,
    timeout: float = DEFAULT_TIMEOUT,
    source_name: str = "",
) -> dict[str, Any]:
    """Ask a local Ollama model to structure invoice lines. No cloud API."""
    snippet = (text or "").strip()
    if not snippet:
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": "",
            "message": "No hay texto para enviar a Ollama.",
        }

    try:
        base = validate_ollama_base(host or default_host())
    except ValueError as exc:
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": "",
            "message": str(exc),
        }

    chosen_model = (model or default_model()).strip()
    chat_url = urljoin(base + "/", "api/chat")
    payload = {
        "model": chosen_model,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0},
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    "Extrae las líneas de producto de esta factura:\n\n"
                    + snippet[: max_prompt_chars(chosen_model)]
                ),
            },
        ],
    }
    try:
        body = _request_json(chat_url, payload=payload, timeout=timeout)
    except HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace")
        except Exception:
            detail = ""
        hint = ""
        if exc.code == 404:
            hint = f" Descargue el modelo con `ollama pull {chosen_model}`."
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": detail,
            "message": f"Ollama HTTP {exc.code} ({chosen_model}).{hint}",
        }
    except URLError:
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": "",
            "message": (
                "No hay conexión con Ollama. Compruebe que el servicio local "
                f"esté activo en {base}."
            ),
        }
    except TimeoutError:
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": "",
            "message": (
                f"Ollama tardó más de {timeout:.0f}s. Use un modelo más pequeño "
                f"como {DEFAULT_MODEL} o aumente el tiempo de espera."
            ),
        }
    except (OSError, json.JSONDecodeError) as exc:
        return {
            "ok": False,
            "rows": [],
            "tax_included": None,
            "raw": "",
            "message": f"Error leyendo la respuesta de Ollama: {exc}",
        }

    content = ""
    message = body.get("message") if isinstance(body, dict) else None
    if isinstance(message, dict):
        content = str(message.get("content") or "")
    if not content and isinstance(body, dict):
        content = str(body.get("response") or "")

    parsed = extract_json_object(content)
    rows = items_from_payload(parsed, source_name=source_name)
    tax_included = parsed.get("tax_included")
    if isinstance(tax_included, str):
        lowered = tax_included.strip().lower()
        if lowered in {"true", "si", "sí", "yes"}:
            tax_included = True
        elif lowered in {"false", "no"}:
            tax_included = False
        else:
            tax_included = None
    elif not isinstance(tax_included, bool):
        tax_included = None

    if not rows:
        return {
            "ok": False,
            "rows": [],
            "tax_included": tax_included,
            "raw": content[:4000],
            "message": (
                f"Ollama ({chosen_model}) no devolvió líneas de producto. "
                "Se mantiene el parser tradicional."
            ),
        }
    return {
        "ok": True,
        "rows": rows,
        "tax_included": tax_included,
        "raw": content[:4000],
        "message": f"Ollama extrajo {len(rows)} línea(s) con {chosen_model}.",
    }
