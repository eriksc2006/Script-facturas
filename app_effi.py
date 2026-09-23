import io
import os
import re
import sys
import json
import math
import sys
import hashlib
import shutil
import logging
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Permitir importar src/effi_processor sin instalación editable.
_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

try:
    from effi_processor.vision_ollama import (
        extract_invoice_rows_from_images,
        ollama_available,
        ollama_vision_model,
    )
except Exception:  # pragma: no cover
    extract_invoice_rows_from_images = None
    ollama_available = None
    ollama_vision_model = lambda: "moondream"  # noqa: E731

# Optional readers/OCR
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from PIL import Image
except Exception:
    Image = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from rapidfuzz import fuzz, process as rf_process
except Exception:
    fuzz = None
    rf_process = None


APP_TITLE = "Effi – Procesador de Facturas e Importación"
PROJECT_ROOT = Path(__file__).resolve().parent
_SRC = PROJECT_ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from effi_processor.ollama_extract import (  # noqa: E402
    check_ollama,
    default_host,
    default_model,
    extract_invoice_items,
    select_extracted_rows,
    should_call_ollama,
)


@st.cache_data(ttl=20, show_spinner=False)
def cached_check_ollama(host: str):
    return check_ollama(host)


LOCAL_TESSDATA = PROJECT_ROOT / "tessdata"
TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    Path("/usr/bin/tesseract"),
    Path("/usr/local/bin/tesseract"),
    Path("/opt/homebrew/bin/tesseract"),
]
TESSERACT_STATE = {"ok": False, "cmd": None, "lang": "eng", "message": ""}


def configure_tesseract():
    """Locate Tesseract binary, local tessdata and OCR language packs."""
    if pytesseract is None:
        TESSERACT_STATE.update(
            ok=False, message="pytesseract no está instalado (pip install pytesseract)."
        )
        return False, TESSERACT_STATE["message"]

    exe = None
    env_cmd = os.environ.get("TESSERACT_CMD") or os.environ.get("TESSERACT_PATH")
    if env_cmd and Path(env_cmd).exists():
        exe = Path(env_cmd)
    else:
        which = shutil.which("tesseract")
        if which:
            exe = Path(which)
        else:
            for candidate in TESSERACT_CANDIDATES:
                if candidate.exists():
                    exe = candidate
                    break

    if exe is None:
        msg = (
            "Tesseract OCR no está instalado o no está en el PATH. "
            "En Windows instale UB-Mannheim Tesseract OCR o defina TESSERACT_CMD."
        )
        TESSERACT_STATE.update(ok=False, message=msg)
        return False, msg

    pytesseract.pytesseract.tesseract_cmd = str(exe)

    # Prefer project tessdata (includes Spanish) over system default.
    if LOCAL_TESSDATA.exists() and any(LOCAL_TESSDATA.glob("*.traineddata")):
        os.environ["TESSDATA_PREFIX"] = str(LOCAL_TESSDATA)

    langs = set()
    try:
        langs = set(pytesseract.get_languages(config=""))
    except Exception:
        for folder in (LOCAL_TESSDATA, exe.parent / "tessdata"):
            if folder.exists():
                langs |= {p.stem for p in folder.glob("*.traineddata")}

    if "spa" in langs and "eng" in langs:
        ocr_lang = "spa+eng"
    elif "spa" in langs:
        ocr_lang = "spa"
    elif "eng" in langs:
        ocr_lang = "eng"
    else:
        ocr_lang = "eng"

    msg = f"Tesseract listo ({exe}, idiomas: {ocr_lang})"
    TESSERACT_STATE.update(ok=True, cmd=str(exe), lang=ocr_lang, message=msg)
    return True, msg


MAIN_SHEET = "Plantilla_Importacion"
AUDIT_SHEET = "Calculos_y_Verificaciones"

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

UNIT_FACTORS_TO_LITERS = {
    "ml": 0.001,
    "mililitro": 0.001,
    "mililitros": 0.001,
    "l": 1.0,
    "lt": 1.0,
    "litro": 1.0,
    "litros": 1.0,
    "cl": 0.01,
    "centilitro": 0.01,
    "centilitros": 0.01,
    "gal": 3.785411784,
    "galon": 3.785411784,
    "galón": 3.785411784,
    "galones": 3.785411784,
    "us gallon": 3.785411784,
}
UNIT_FACTORS_TO_KG = {
    "mg": 0.000001,
    "miligramo": 0.000001,
    "miligramo": 0.000001,
    "g": 0.001,
    "gramo": 0.001,
    "gramos": 0.001,
    "kg": 1.0,
    "kilo": 1.0,
    "kilos": 1.0,
    "kilogramo": 1.0,
    "kilogramos": 1.0,
}
UNIT_FACTORS_TO_UNITS = {"und": 1.0, "unidad": 1.0, "unidades": 1.0, "u": 1.0}

MONEY_RE = r"[-+]?\$?\s*[\d\.,]+"
PCT_RE = r"(\d+(?:[\.,]\d+)?)\s*%"
EFFI_MAX_UPLOAD_BYTES = 5 * 1024 * 1024  # 5 MB (límite Effi)
EFFI_IMPORT_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlt"}
INVOICE_UPLOAD_TYPES = [
    "pdf",
    "png",
    "jpg",
    "jpeg",
    "webp",
    "bmp",
    "tif",
    "tiff",
    "csv",
    "xlsx",
    "xls",
    "xlsm",
    "xlt",
    "txt",
]


def normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parse_number(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (list, tuple)):
        # OCR / findall a veces deja listas en celdas; tomar el primer número usable.
        for item in value:
            parsed = parse_number(item)
            if parsed is not None:
                return parsed
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace("%", "").replace(" ", "")
    if not s:
        return None
    # Colombian / European formats: 1.234.567,89
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # decimal comma unless many groups imply thousands
        parts = s.split(",")
        s = (
            "".join(parts[:-1]) + "." + parts[-1]
            if len(parts[-1]) <= 2
            else "".join(parts)
        )
    elif "." in s:
        parts = s.split(".")
        s = "".join(parts) if len(parts[-1]) == 3 and len(parts) > 1 else s
    try:
        return float(s)
    except Exception:
        return None


def excel_safe_value(value):
    """openpyxl solo acepta escalares; listas/dicts rompen la exportación."""
    if value is None:
        return None
    if isinstance(value, (list, tuple, set)):
        return " | ".join(
            excel_safe_value(v) if not isinstance(v, str) else v for v in value
        )
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, float) and math.isnan(value):
        return None
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        try:
            return value.item()
        except Exception:
            pass
    return value


def parse_money_series(series):
    return series.map(parse_number)


def _ensure_ocr_ready():
    ok, msg = configure_tesseract()
    if not ok:
        raise RuntimeError(msg)
    if Image is None:
        raise RuntimeError("Para OCR instale Pillow: pip install Pillow")
    return TESSERACT_STATE["lang"]


def ocr_pil_image(img):
    lang = _ensure_ocr_ready()
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    try:
        return pytesseract.image_to_string(img, lang=lang)
    except pytesseract.TesseractError:
        # Fallback if spa+eng fails for any reason
        return pytesseract.image_to_string(img, lang="eng")


def extract_text_from_pdf(file_bytes):
    if fitz is None:
        raise RuntimeError("Instale PyMuPDF para procesar PDF: pip install pymupdf")
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages = []
    plain_chars = 0
    for i, page in enumerate(doc):
        txt = page.get_text("text") or ""
        plain_chars += len(txt.strip())
        pages.append(f"\n--- PÁGINA {i+1} ---\n{txt}")

    # Scanned / image-only PDF: run OCR page by page.
    if plain_chars < 40:
        ocr_pages = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            ocr_txt = ocr_pil_image(img)
            ocr_pages.append(f"\n--- PÁGINA {i+1} (OCR) ---\n{ocr_txt}")
        return "\n".join(ocr_pages), "pdf-ocr"
    return "\n".join(pages), "pdf"


def extract_text_from_image(file_bytes):
    img = Image.open(io.BytesIO(file_bytes))
    return ocr_pil_image(img)


def pdf_plain_text_probe(file_bytes):
    """Texto nativo del PDF sin OCR (para decidir Vision vs texto)."""
    if fitz is None:
        raise RuntimeError("Instale PyMuPDF para procesar PDF: pip install pymupdf")
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    parts = []
    for i, page in enumerate(doc):
        parts.append(page.get_text("text") or "")
    text = "\n".join(parts)
    return text, len(text.strip())


def pdf_page_images(file_bytes, max_pages: int = 8):
    """Renderiza páginas PDF a imágenes PIL (para Vision / OCR)."""
    if fitz is None:
        raise RuntimeError("Instale PyMuPDF para procesar PDF: pip install pymupdf")
    if Image is None:
        raise RuntimeError("Instale Pillow: pip install Pillow")
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    images = []
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        images.append(Image.open(io.BytesIO(pix.tobytes("png"))))
    return images


def image_file_to_pil(file_bytes):
    if Image is None:
        raise RuntimeError("Instale Pillow: pip install Pillow")
    return Image.open(io.BytesIO(file_bytes))


def check_ollama_vision(model=None):
    if ollama_available is None:
        return False, "Módulo Vision no disponible."
    return ollama_available(model=model)


def try_vision_extract(images, source_name: str, model=None):
    """Intenta extracción Vision; lanza si falla."""
    if extract_invoice_rows_from_images is None:
        raise RuntimeError("Módulo Vision Ollama no importado.")
    ok, msg = check_ollama_vision(model=model)
    if not ok:
        raise RuntimeError(msg)
    return extract_invoice_rows_from_images(
        images, source_name=source_name, model=model
    )


def read_document(uploaded_file):
    ext = Path(uploaded_file.name).suffix.lower()
    data = uploaded_file.getvalue()
    if ext == ".pdf":
        text, kind = extract_text_from_pdf(data)
        return text, kind
    if ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        return extract_text_from_image(data), "ocr"
    if ext in {".txt", ".csv"}:
        return data.decode("utf-8", errors="replace"), "text"
    if ext in {".xlsx", ".xls", ".xlsm", ".xlt"}:
        return None, "spreadsheet"
    raise ValueError(f"Formato no soportado: {ext}")


def read_catalog(uploaded_file):
    data = uploaded_file.getvalue()
    ext = Path(uploaded_file.name).suffix.lower()
    if ext == ".csv":
        try:
            return pd.read_csv(io.BytesIO(data), dtype=str)
        except Exception:
            return pd.read_csv(io.BytesIO(data), sep=";", dtype=str)
    if ext in {".xlsx", ".xls", ".xlsm", ".xlt"}:
        return pd.read_excel(io.BytesIO(data), dtype=str)
    raise ValueError("El catálogo maestro debe ser CSV o Excel (xlsx|xlsm|xls|xlt).")


def find_column(df, candidates, *, exclude_substrings=None):
    """
    Busca columna por nombre.
    1) coincidencia exacta normalizada
    2) el candidato aparece dentro del nombre de columna (no al revés:
       evita que 'ID' matchee 'Unidad' o 'Código' matchee 'Código Impuesto').
    """
    exclude_substrings = [normalize_text(x) for x in (exclude_substrings or [])]
    norm_map = {normalize_text(c): c for c in df.columns}

    def _excluded(nc: str) -> bool:
        return any(ex and ex in nc for ex in exclude_substrings)

    for candidate in candidates:
        n = normalize_text(candidate)
        if n in norm_map and not _excluded(n):
            return norm_map[n]

    for candidate in candidates:
        n = normalize_text(candidate)
        if len(n) < 3:
            continue
        for c in df.columns:
            nc = normalize_text(c)
            if _excluded(nc):
                continue
            if n in nc:
                return c
    return None


# Mapeo tarifa → Código Effi Impuesto (Colombia / plantillas Effi del proyecto).
# 19% → 1 (confirmado en plantillas corregidas del usuario). 0% → vacío.
DEFAULT_EFFI_TAX_BY_RATE = {
    0.0: "",
    0.05: "2",
    0.19: "1",
}


def detect_catalog_columns(df):
    return {
        "gtin": find_column(df, [
            "COD. BARRAS GTIN", "Código de barras GTIN", "Codigo de barras GTIN",
            "GTIN", "Código de barras", "Codigo de barras", "EAN", "Barcode",
            "COD. BARRAS",
        ]),
        "reference": find_column(df, [
            "Código", "Codigo", "Código producto", "Codigo producto",
            "Referencia", "Ref", "SKU", "Código interno", "Codigo interno",
        ]),
        "effi": find_column(df, [
            "ID EFFI", "Código EFFI", "Codigo EFFI", "ID Artículo", "Articulo ID", "ID",
        ]),
        "description": find_column(df, [
            "Descripción", "Descripcion", "Nombre", "Artículo", "Articulo", "Producto",
        ]),
        "presentation": find_column(df, [
            "Presentación", "Presentacion", "Empaque", "Contenido", "Tamaño", "Tamano", "Unidad",
        ]),
        "brand": find_column(df, ["Marca", "Brand"]),
        # Solo columnas de CÓDIGO de impuesto; nunca precios con la palabra impuesto/IVA.
        "tax": find_column(df, [
            "Código Effi Impuesto", "Codigo Effi Impuesto",
            "Código Impuesto", "Codigo Impuesto",
            "ID Impuesto", "Impuesto ID", "ID IVA",
        ], exclude_substrings=["precio", "tarifa", "utilidad", "valor"]),
        "price_net": find_column(df, [
            "Precio: TARIFA NORMAL", "Precio TARIFA NORMAL",
            "Precio neto", "Precio sin impuesto", "Precio sin IVA",
        ], exclude_substrings=["impuesto", "iva", "utilidad", "+"]),
        "price_gross": find_column(df, [
            "Precio + impuesto: TARIFA NORMAL",
            "Precio + impuesto TARIFA NORMAL",
            "Precio + impuesto",
            "Precio con impuesto",
            "Precio con IVA",
            "Precio IVA incluido",
        ], exclude_substrings=["utilidad", "mayorista"]),
    }


def looks_like_effi_tax_code(value):
    """Códigos Effi de impuesto son IDs cortos (ej. 1), no precios ni tarifas."""
    if value is None:
        return False
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none"}:
        return False
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    if not re.fullmatch(r"\d{1,4}", s):
        return False
    return 1 <= int(s) <= 9999


def infer_tax_rate_from_catalog_prices(net_value, gross_value):
    net = parse_number(net_value)
    gross = parse_number(gross_value)
    if net is None or gross is None or net <= 0 or gross < 0:
        return None
    if abs(gross - net) < 1e-6:
        return 0.0
    rate = (gross / net) - 1.0
    if rate < -0.01 or rate > 1.0:
        return None
    # Normalizar a tarifas colombianas conocidas
    for known in (0.0, 0.05, 0.19):
        if abs(rate - known) <= 0.015:
            return known
    return round(rate, 4)


def map_rate_to_effi_tax_code(rate, tax_by_rate=None, default_code="1"):
    mapping = dict(DEFAULT_EFFI_TAX_BY_RATE)
    if tax_by_rate:
        mapping.update(tax_by_rate)
    if rate is None:
        return str(default_code)
    best_code = None
    best_diff = 1e9
    for known_rate, code in mapping.items():
        diff = abs(float(known_rate) - float(rate))
        if diff < best_diff:
            best_diff = diff
            best_code = code
    if best_diff <= 0.02:
        return "" if best_code is None else str(best_code)
    return str(default_code)


def resolve_effi_tax_code(catalog_row, cols, settings=None):
    """
    Obtiene el Código Effi Impuesto desde el maestro Effi:
    1) columna explícita de código (si es un ID válido)
    2) tarifa inferida de Precio vs Precio+impuesto del catálogo
    3) fallback al IVA configurado en la UI
    """
    settings = settings or {}
    default_code = str(settings.get("default_tax_code", "1"))
    tax_by_rate = settings.get("tax_by_rate")

    if cols.get("tax"):
        raw = catalog_row.get(cols["tax"], "")
        if looks_like_effi_tax_code(raw):
            s = str(raw).strip()
            if re.fullmatch(r"\d+\.0+", s):
                s = s.split(".", 1)[0]
            return s, "columna código impuesto Effi", None

    rate = None
    source = "fallback UI"
    if cols.get("price_net") and cols.get("price_gross"):
        rate = infer_tax_rate_from_catalog_prices(
            catalog_row.get(cols["price_net"]),
            catalog_row.get(cols["price_gross"]),
        )
        if rate is not None:
            source = "precio vs precio+impuesto Effi"

    if rate is None:
        rate = float(settings.get("tax_rate", 0.19) or 0.19)
        source = "IVA configurado UI"

    code = map_rate_to_effi_tax_code(rate, tax_by_rate=tax_by_rate, default_code=default_code)
    return code, source, rate


def clean_gtin(value):
    if value is None:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s


def clean_product_code(value):
    """Preserve Effi identifiers, including alphanumeric references like SILMANTSH."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        return clean_gtin(text)
    return re.sub(r"[^A-Za-z0-9._/-]", "", text).upper()


def extract_presentation(text):
    s = normalize_text(text)
    # e.g. 3.785 L, 4L, 500 ml, 1.5 kg, 250 g, 12 und
    patterns = [
        (
            r"(\d+(?:\.\d+)?)\s*(ml|mililitros?|cl|litros?|l|lt|galones?|galon|gal)\b",
            "volume",
        ),
        (r"(\d+(?:\.\d+)?)\s*(mg|gramos?|g|kg|kilos?)\b", "mass"),
        (r"(\d+(?:\.\d+)?)\s*(und|unidades?|unidad|u)\b", "units"),
    ]
    for pat, kind in patterns:
        m = re.search(pat, s)
        if m:
            return float(m.group(1)), m.group(2), kind
    # "galon de 3.785L" / "galón 3,785 litros"
    if "galon" in s or "galones" in s:
        m = re.search(r"(\d+(?:\.\d+)?)\s*l\b", s)
        if m:
            return float(m.group(1)), "l", "volume"
        return 1.0, "galon", "volume"
    return None, None, None


def presentation_base_quantity(text):
    n, unit, kind = extract_presentation(text)
    if n is None:
        return None, None
    u = normalize_text(unit)
    if kind == "volume":
        factor = UNIT_FACTORS_TO_LITERS.get(u)
    elif kind == "mass":
        factor = UNIT_FACTORS_TO_KG.get(u)
    else:
        factor = UNIT_FACTORS_TO_UNITS.get(u)
    return (n * factor if factor else n), kind


def similarity(a, b):
    a, b = normalize_text(a), normalize_text(b)
    if not a or not b:
        return 0
    if fuzz:
        return fuzz.token_set_ratio(a, b) / 100.0
    # fallback Jaccard
    A, B = set(a.split()), set(b.split())
    return len(A & B) / max(1, len(A | B))


def build_catalog_index(catalog, cols):
    """Índice GTIN/código alfanumérico + textos precomputados."""
    gtin_map = {}
    if cols.get("gtin"):
        for idx, raw in catalog[cols["gtin"]].items():
            gtin = clean_gtin(raw)
            if gtin and gtin not in gtin_map:
                gtin_map[gtin] = idx

    reference_map = {}
    if cols.get("reference"):
        for idx, raw in catalog[cols["reference"]].items():
            reference = clean_product_code(raw)
            if reference and reference not in reference_map:
                reference_map[reference] = idx

    texts = []
    indices = []
    for idx, row in catalog.iterrows():
        parts = []
        for k in ("description", "presentation", "brand"):
            col = cols.get(k)
            if col:
                parts.append(str(row.get(col, "") or ""))
        text = normalize_text(" ".join(parts))
        if text:
            texts.append(text)
            indices.append(idx)

    return {
        "gtin_map": gtin_map,
        "reference_map": reference_map,
        "texts": texts,
        "indices": indices,
    }


def _token_overlap_score(a: str, b: str) -> float:
    sa = set(normalize_text(a).split())
    sb = set(normalize_text(b).split())
    if not sa or not sb:
        return 0.0
    overlap = len(sa & sb)
    union = len(sa | sb)
    return overlap / max(1, union)


def match_catalog_item(description, presentation, catalog, cols, catalog_index=None):
    query = f"{description or ''} {presentation or ''}".strip()
    qn = normalize_text(query)
    if not qn:
        return None, 0, "Sin descripción"

    qgt = clean_gtin(description)
    if qgt and catalog_index and qgt in catalog_index["gtin_map"]:
        idx = catalog_index["gtin_map"][qgt]
        return catalog.loc[idx], 1.0, "GTIN en descripción"

    query_tokens = set(qn.split())
    best_row = None
    best_score = 0.0
    best_method = "Fuzzy"

    if catalog_index and catalog_index["texts"]:
        texts = catalog_index["texts"]
        indices = catalog_index["indices"]

        for pos, text in enumerate(texts):
            idx = indices[pos]
            text_tokens = set(text.split())
            token_overlap = len(query_tokens & text_tokens) / max(1, len(query_tokens | text_tokens))
            sim = similarity(qn, text)
            score = max(sim, token_overlap * 0.9)

            # Penalizar si el texto parece una firma/encabezado o un campo genérico.
            if re.search(r"\b(factura|subtotal|total|iva|cliente|nit|telefono|direccion)\b", text):
                score *= 0.4

            if score > best_score:
                best_score = score
                best_row = catalog.loc[idx]
                best_method = "token" if token_overlap > sim else "Fuzzy"

        if best_row is not None and best_score >= 0.35:
            return best_row, best_score, best_method

        if rf_process and fuzz:
            result = rf_process.extractOne(
                qn,
                texts,
                scorer=fuzz.token_set_ratio,
                score_cutoff=20,
            )
            if result is not None:
                _match, score100, pos = result
                return catalog.loc[indices[pos]], score100 / 100.0, "Fuzzy"

    # Fallback legacy con búsqueda más flexible por nombre y presentación.
    for idx, row in catalog.iterrows():
        text = " ".join(
            str(row.get(cols[k], "") or "")
            for k in ["description", "presentation", "brand"]
            if cols.get(k)
        )
        score = similarity(query, text)
        token_overlap = _token_overlap_score(query, text)
        score = max(score, token_overlap * 0.9)

        if qgt and cols.get("gtin"):
            cgt = clean_gtin(row.get(cols["gtin"], ""))
            if cgt and qgt == cgt:
                score = 1.0

        if score > best_score:
            best_score = score
            best_row = row
            best_method = "token" if token_overlap > 0.35 else "Fuzzy"

    if best_row is None:
        return None, 0, "Sin coincidencia"
    return best_row, best_score, best_method


def infer_columns_from_invoice_df(df):
    return {
        "code": find_column(
            df,
            [
                "Referencia",
                "REF",
                "Codigo",
                "Código",
                "GTIN",
                "EAN",
                "Barcode",
                "Código de barras",
                "SKU",
            ],
        ),
        "description": find_column(
            df,
            [
                "Descripcion",
                "Descripción",
                "Articulo",
                "Artículo",
                "Producto",
                "Nombre",
            ],
        ),
        "presentation": find_column(
            df, ["Presentacion", "Presentación", "Empaque", "Unidad"]
        ),
        "info": find_column(df, ["Inf.", "Inf", "Información", "Info"]),
        "quantity": find_column(df, ["Cantidad", "Cant.", "Cant"]),
        "list_price": find_column(
            df, ["Precio Lista", "Precio de lista", "Precio Lista Unitario", "P. Lista"]
        ),
        "unit_price": find_column(
            df,
            [
                "Precio Unitario",
                "PRECIO UNITARIO",
                "Precio ud.",
                "Precio",
                "Valor Unitario",
            ],
        ),
        "total": find_column(
            df, ["Valor", "VALOR", "Precio Total", "Total", "Valor Total", "Importe"]
        ),
        "discount": find_column(
            df, ["Descuento", "Desc.", "Dto.", "% Descuento", "% Desc"]
        ),
        "tax": find_column(df, ["IVA", "Impuesto", "Tax"]),
    }


def extract_invoice_rows_from_text(text):
    """
    Heurística para documentos de texto/OCR.
    Une líneas partidas por OCR y filtra encabezados/pies de factura.
    """
    noise_re = re.compile(
        r"\b("
        r"factura|nit\b|rut\b|regimen|resolucion|autorizacion|cufe|cuds|"
        r"pagina|page\b|fecha|hora|cliente|vendedor|forma de pago|"
        r"subtotal|total a pagar|valor total|gran total|total factura|"
        r"retefuente|reteiva|reteica|impuesto|iva\b|base gravable|"
        r"direccion|telefono|correo|email|www\.|http|"
        r"cantidad|descripcion|producto|codigo|precio|descuento|valor|"
        r"observaciones|gracias|firma|representante"
        r")\b",
        re.I,
    )
    # Pure header row with many column titles and few product-like tokens
    headerish_re = re.compile(
        r"(codigo|descripcion|cantidad).*(precio|valor|total)",
        re.I,
    )

    raw_lines = [ln.strip() for ln in (text or "").splitlines()]
    # Merge a text-only line with the next numeric line (common OCR split).
    merged = []
    i = 0
    while i < len(raw_lines):
        line = raw_lines[i]
        if not line:
            i += 1
            continue
        nums_here = re.findall(MONEY_RE, line)
        if (
            i + 1 < len(raw_lines)
            and len(nums_here) <= 1
            and re.search(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]", line)
            and re.findall(MONEY_RE, raw_lines[i + 1] or "")
            and len(re.sub(MONEY_RE, "", raw_lines[i + 1]).strip()) < 8
        ):
            merged.append(f"{line} {raw_lines[i + 1].strip()}")
            i += 2
            continue
        merged.append(line)
        i += 1

    rows = []
    for line_no, line in enumerate(merged, 1):
        n = normalize_text(line)
        if not n or len(n) < 3:
            continue
        if headerish_re.search(line) and len(re.findall(MONEY_RE, line)) < 2:
            continue
        nums = re.findall(MONEY_RE, line)
        if not nums:
            continue
        # Skip short noise / totals without product description length
        alpha_len = len(re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúñÑ]", "", line))
        if (
            noise_re.search(n)
            and len(nums) < 2
            and alpha_len < 18
            and not re.match(r"^\*?\d{6,}", line.strip())
        ):
            continue
        # Need some alphabetic description OR a product code.
        if alpha_len < 3 and not re.search(r"\d{6,}", line):
            continue
        rows.append({"linea": line_no, "texto": line, "numeros": nums})
    return rows


def _number_spans(line):
    """Return [(start, end, raw, parsed), ...] ignoring presentation sizes (4L, 3.785L, 200g)."""
    unit_suffix = (
        r"(?:ml|mililitros?|cl|litros?|l|lt|galones?|galon|gal|"
        r"mg|gramos?|g|kg|kilos?|kilogramos?|"
        r"und|unidades?|unidad|u)\b"
    )
    spans = []
    for m in re.finditer(MONEY_RE, line):
        raw = m.group(0)
        tail = line[m.end() : m.end() + 12]
        if re.match(rf"^\s*{unit_suffix}", tail, flags=re.I):
            continue
        parsed = parse_number(raw)
        if parsed is None:
            continue
        spans.append((m.start(), m.end(), raw, float(parsed)))
    return spans


def parse_text_line(line):
    """
    Parser flexible para líneas OCR/PDF:
    - Formato THYM'S / típico CO: REF CANT DESC PRECIO VALOR
    - Formato genérico: [ * ] [CODIGO] DESCRIPCION ... CANT PRECIO TOTAL [%DESC]
    """
    if not line or not str(line).strip():
        return None
    original = str(line).strip()
    work = original

    is_bonus = bool(re.match(r"^\*", work)) or bool(
        re.search(r"\b(bonific|obsequi|regalo|\bx\s*y\b)\b", normalize_text(work))
    )
    # Asterisco manuscrito/OCR pegado a la referencia: TA91* / *TA91
    if re.search(r"(^\*|[A-Za-z0-9]\s*\*(?=\s|$))", work):
        is_bonus = True
    work = work.lstrip("*").strip()
    work = re.sub(r"(?<=[A-Za-z0-9])\*(?=\s|$)", "", work).strip()

    disc_pct = None
    pct_m = re.search(r"(\d+(?:[.,]\d+)?)\s*%\s*$", work)
    if pct_m:
        disc_pct = float(pct_m.group(1).replace(",", "."))
        work = work[: pct_m.start()].strip()

    # --- Formato factura tipo THYM'S: REF  CANT  DESCRIPCION  PRECIO  VALOR ---
    thyms = re.match(
        r"^([A-Za-z0-9][A-Za-z0-9.\-/]{0,24})\s+"
        r"(\d+(?:[.,]\d+)?)\s+"
        r"(.+?)\s+"
        r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})|\d+[.,]\d{2}|\d+)\s+"
        r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})|\d+[.,]\d{2}|\d+)\s*$",
        work,
    )
    if thyms:
        code_t, qty_s, desc_t, unit_s, total_s = thyms.groups()
        if re.search(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]", desc_t):
            qty_t = parse_number(qty_s)
            unit_t = parse_number(unit_s)
            total_t = parse_number(total_s)
            # Validar coherencia qty * precio ≈ valor (tolerancia 15%)
            coherent = True
            if qty_t and unit_t and total_t is not None and qty_t > 0:
                expected = qty_t * unit_t
                if abs(expected - total_t) > max(1.0, 0.15 * abs(total_t)):
                    coherent = False
            if coherent and qty_t is not None:
                n_pres, u_pres, _kind = extract_presentation(desc_t)
                presentation = f"{n_pres} {u_pres}" if n_pres is not None else ""
                return {
                    "code": code_t.strip(),
                    "description": re.sub(r"\s+", " ", desc_t).strip(),
                    "presentation": presentation,
                    "quantity": qty_t,
                    "unit_price_invoice": unit_t,
                    "total_invoice": total_t,
                    "discount_pct": disc_pct,
                    "discount_value": None,
                    "is_bonus": is_bonus,
                    "source_line": original,
                }

    code = ""
    code_match = re.match(r"^(\d{6,14})\b", work)
    if code_match:
        code = clean_gtin(code_match.group(1))
        work = work[code_match.end() :].strip()
    else:
        alt = re.match(r"^([A-Z0-9][A-Z0-9.\-/]{1,24})\b", work, re.I)
        if alt and not parse_number(alt.group(1)):
            token = alt.group(1)
            # Referencias cortas tipo VT20, T1.0, TA91, TP124
            if re.search(r"\d", token) or "-" in token or "/" in token or "." in token:
                code = token.strip()
                work = work[alt.end() :].strip()

    spans = _number_spans(work)
    if not spans:
        return None

    trailing = []
    for sp in reversed(spans):
        if not trailing:
            trailing.append(sp)
            continue
        prev = trailing[-1]
        gap = prev[0] - sp[1]
        if gap <= 8 and len(trailing) < 4:
            trailing.append(sp)
        else:
            break
    trailing = list(reversed(trailing))

    cut = trailing[0][0]
    desc = work[:cut].strip(" -|:\t")
    desc = re.sub(r"\s+", " ", desc).strip()

    # Si la descripción empieza con cantidad aislada (REF ya extraído): "36 TRATAMIENTO..."
    qty_prefix = re.match(r"^(\d+(?:[.,]\d+)?)\s+(.+)$", desc)
    leading_qty = None
    if qty_prefix and re.search(r"[A-Za-zÁÉÍÓÚáéíóúñÑ]", qty_prefix.group(2)):
        leading_qty = parse_number(qty_prefix.group(1))
        desc = qty_prefix.group(2).strip()

    if not desc and code:
        desc = code
    if not desc:
        return None

    nd = normalize_text(desc)
    if nd in {"total", "subtotal", "iva", "descuento", "neto", "bruto"}:
        return None

    vals = [sp[3] for sp in trailing]
    qty = unit = total = disc_val = None

    def looks_qty(v):
        return v is not None and abs(v - round(v)) < 1e-9 and 0 < v <= 50000

    def looks_money(v):
        return v is not None and v >= 0

    if leading_qty is not None and len(vals) >= 2:
        qty = leading_qty
        if len(vals) >= 2:
            unit, total = vals[-2], vals[-1]
        else:
            total = vals[-1]
    elif len(vals) == 1:
        total = vals[0]
        qty = 1.0
    elif len(vals) == 2:
        a, b = vals
        if looks_qty(a) and looks_money(b) and (b >= a or a <= 1000):
            qty, total = a, b
            unit = (b / a) if a else None
        else:
            unit, total = a, b
            qty = 1.0
    else:
        a, b, c = vals[0], vals[1], vals[2]
        if looks_qty(a) and looks_money(b) and looks_money(c):
            qty, unit, total = a, b, c
            if (
                qty
                and unit
                and total
                and abs(qty * unit - total) > max(1.0, 0.25 * abs(total))
            ):
                if a > 100:
                    unit, total = a, b
                    qty = (total / unit) if unit else 1.0
                    if disc_pct is None and c <= 100:
                        disc_pct = c
        else:
            qty, unit, total = a, b, c
        if len(vals) >= 4 and disc_pct is None:
            if vals[3] <= 100:
                disc_pct = vals[3]
            else:
                disc_val = vals[3]

    if is_bonus and (total is None or total == 0):
        disc_pct = 100.0 if disc_pct is None else disc_pct

    n_pres, u_pres, _kind = extract_presentation(desc)
    presentation = f"{n_pres} {u_pres}" if n_pres is not None else ""

    return {
        "code": code,
        "description": desc,
        "presentation": presentation,
        "quantity": qty,
        "unit_price_invoice": unit,
        "total_invoice": total,
        "discount_pct": disc_pct,
        "discount_value": disc_val,
        "is_bonus": is_bonus,
        "source_line": original,
    }


def parse_invoice_text(text):
    """Parse full invoice text into structured item rows."""
    parsed = []
    for item in extract_invoice_rows_from_text(text or ""):
        row = parse_text_line(item["texto"])
        if row:
            parsed.append(row)
    return parsed


def detect_global_tax(text):
    n = normalize_text(text)
    if "iva incluido" in n or "iva incluido" in n:
        return True
    return None


def find_percentages(text):
    return [float(x.replace(",", ".")) for x in re.findall(PCT_RE, text)]


def find_additional_charges(text):
    keywords = [
        "retefuente",
        "rete iva",
        "reteiva",
        "rete ica",
        "reteica",
        "impuesto al consumo",
        "flete",
        "transporte",
        "seguro",
        "otros cargos",
        "descuento financiero",
        "descuento pronto pago",
    ]
    out = []
    lines = text.splitlines()
    for line in lines:
        n = normalize_text(line)
        if any(k in n for k in keywords):
            vals = re.findall(MONEY_RE, line)
            # Unir a texto: openpyxl no acepta listas en celdas.
            out.append(
                {
                    "concepto": line.strip(),
                    "valores_detectados": (
                        " | ".join(v.strip() for v in vals) if vals else ""
                    ),
                }
            )
    return out


def detect_bonus(text):
    """
    Detecta asteriscos/bonificaciones y patrones como 10 + 2 gratis.
    """
    bonuses = []
    for line in text.splitlines():
        n = normalize_text(line)
        if "*" in line or any(
            k in n for k in ["bonificacion", "obsequio", "gratis", "regalo", "bono"]
        ):
            bonuses.append(line.strip())
    return bonuses


def compute_invoice_item(row, catalog_row, cols, tax_rate=0.19):
    qty = parse_number(row.get("quantity"))
    unit_invoice = parse_number(row.get("unit_price_invoice"))
    total_invoice = parse_number(row.get("total_invoice"))
    discount_pct = parse_number(row.get("discount_pct"))
    discount_value = parse_number(row.get("discount_value"))

    qty = qty if qty is not None else 0.0

    if total_invoice is None and unit_invoice is not None:
        total_invoice = qty * unit_invoice
    if unit_invoice is None and total_invoice is not None and qty:
        unit_invoice = total_invoice / qty
    if unit_invoice is None:
        unit_invoice = 0.0
    if total_invoice is None:
        total_invoice = qty * unit_invoice

    # Interpret invoice price as tax-inclusive by default only when global evidence says so.
    # The UI exposes the tax mode, so this function receives the selected mode below.
    return {
        "qty": qty,
        "unit_invoice": unit_invoice,
        "total_invoice": total_invoice,
        "discount_pct": discount_pct or 0.0,
        "discount_value": discount_value or 0.0,
    }


def make_workbook(main_rows, audit_tables):
    wb = openpyxl_workbook()
    ws = wb.active
    ws.title = MAIN_SHEET

    # MAIN SHEET: values/text only; no formulas.
    for c, h in enumerate(EFFI_HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    for r_idx, row in enumerate(main_rows, 2):
        for c_idx, value in enumerate(row, 1):
            value = excel_safe_value(value)
            # Force GTIN and all identifiers to plain text.
            if c_idx == 1 and value is not None:
                value = str(value)
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.number_format = "@"
            else:
                ws.cell(row=r_idx, column=c_idx, value=value)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{max(1, ws.max_row)}"
    widths = [36, 16, 14, 14, 42, 15, 18, 24, 22]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 45

    # Secondary audit sheet: formulas are allowed here, but we write values by default.
    for sheet_name, df in audit_tables.items():
        sh = wb.create_sheet(sheet_name[:31])
        for c, col in enumerate(df.columns, 1):
            sh.cell(1, c, col).font = Font(bold=True)
            sh.cell(1, c).alignment = Alignment(wrap_text=True, vertical="center")
        for r, values in enumerate(df.itertuples(index=False, name=None), 2):
            for c, value in enumerate(values, 1):
                sh.cell(r, c, excel_safe_value(value))
        for c in range(1, sh.max_column + 1):
            sh.column_dimensions[get_column_letter(c)].width = min(
                45, max(14, len(str(sh.cell(1, c).value or "")) + 2)
            )
        sh.freeze_panes = "A2"

    return wb


def openpyxl_workbook():
    from openpyxl import Workbook

    return Workbook()


def save_workbook(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    data = buf.getvalue()
    if len(data) > EFFI_MAX_UPLOAD_BYTES:
        raise ValueError(
            f"El Excel generado pesa {len(data) / (1024 * 1024):.2f} MB; "
            f"Effi admite máximo 5 MB (formatos: xlsx|xlsm|xls|xlt)."
        )
    return data


def spreadsheet_rows_from_df(invoice_df_raw, source_name=""):
    inv_cols = infer_columns_from_invoice_df(invoice_df_raw)
    normalized_rows = []
    for _, r in invoice_df_raw.iterrows():
        code_raw = r.get(inv_cols["code"], "") if inv_cols["code"] else ""
        is_bonus = "*" in str(code_raw) or "*" in " ".join(str(x) for x in r.tolist())
        normalized_rows.append(
            {
                "archivo_origen": source_name,
                "code": str(code_raw).replace("*", "").strip(),
                "description": (
                    r.get(inv_cols["description"], "")
                    if inv_cols["description"]
                    else ""
                ),
                "presentation": (
                    r.get(inv_cols["presentation"], "")
                    if inv_cols["presentation"]
                    else ""
                ),
                "quantity": (
                    r.get(inv_cols["quantity"], "") if inv_cols["quantity"] else ""
                ),
                "unit_price_invoice": (
                    r.get(inv_cols["unit_price"], "")
                    if inv_cols["unit_price"]
                    else r.get(inv_cols["list_price"], "")
                ),
                "total_invoice": (
                    r.get(inv_cols["total"], "") if inv_cols["total"] else ""
                ),
                "discount_pct": (
                    r.get(inv_cols["discount"], "") if inv_cols["discount"] else ""
                ),
                "discount_value": "",
                "is_bonus": is_bonus,
                "source_line": " | ".join(
                    str(x) for x in r.tolist() if str(x) != "nan"
                ),
            }
        )
    return normalized_rows, inv_cols


def load_invoice_uploaded_file(
    uploaded_file,
    ollama_settings=None,
    read_mode: str = "Auto (Vision si hay Ollama)",
):
    """Lee una factura (PDF/imagen/Excel/CSV/TXT) y devuelve filas normalizadas + metadatos."""
    name = uploaded_file.name
    data = uploaded_file.getvalue()
    size = len(data)
    ext = Path(name).suffix.lower()

    # Validar tamaño solo para archivos Excel de entrada tipo plantilla Effi
    if ext in EFFI_IMPORT_EXTENSIONS and size > EFFI_MAX_UPLOAD_BYTES:
        raise ValueError(
            f"'{name}' supera 5 MB ({size / (1024 * 1024):.2f} MB). "
            "Effi limita la importación a 5 MB."
        )

    image_exts = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
    use_vision = read_mode in {
        "Auto (Vision si hay Ollama)",
        "Solo Vision",
    }
    force_tesseract = read_mode == "Solo Tesseract"
    vision_only = read_mode == "Solo Vision"

    # Spreadsheet / texto plano: sin Vision
    if ext in {".xlsx", ".xls", ".xlsm", ".xlt"}:
        raw = pd.read_excel(io.BytesIO(data), dtype=str)
        rows, inv_cols = spreadsheet_rows_from_df(raw, source_name=name)
        return {
            "name": name,
            "kind": "spreadsheet",
            "text": "",
            "rows": rows,
            "candidates": [],
            "columns": inv_cols,
            "ollama_message": None,
            "ollama_tax_included": None,
            "error": None,
        }
    if ext in {".txt", ".csv"}:
        # Continúa hacia la selección de Ollama para priorizarlo también en TXT/CSV.
        invoice_text = data.decode("utf-8", errors="replace")

    # PDF / imagen: Vision opcional + fallback Tesseract
    images = None
    vision_error = None
    if use_vision and not force_tesseract and ext in (image_exts | {".pdf"}):
        try:
            if ext == ".pdf":
                if vision_only:
                    images = pdf_page_images(data)
                else:
                    text_probe, plain_chars = pdf_plain_text_probe(data)
                    # Aunque el PDF tenga texto, se intenta Vision primero para
                    # conservar la estructura de la tabla; el texto queda de respaldo.
                    images = pdf_page_images(data)
            else:
                images = [image_file_to_pil(data)]

            rows, vision_text = try_vision_extract(
                images,
                source_name=name,
                model=(ollama_settings or {}).get("model"),
            )
            if rows:
                return {
                    "name": name,
                    "kind": "vision-ollama",
                    "text": vision_text,
                    "rows": rows,
                    "candidates": [],
                    "columns": None,
                    "error": None,
                }
            vision_error = "Vision no devolvió ítems estructurados."
            if vision_only:
                raise RuntimeError(vision_error)
        except Exception as exc:
            vision_error = str(exc)
            if vision_only:
                raise RuntimeError(
                    f"Vision LLM falló para '{name}': {vision_error}"
                ) from exc

    # Fallback Tesseract / texto PDF
    if ext in {".txt", ".csv"}:
        invoice_kind = "text"
    else:
        invoice_text, invoice_kind = read_document(uploaded_file)
    if vision_error:
        invoice_text = (
            f"[Vision no usado/falló: {vision_error}]\n\n{invoice_text or ''}"
        )
        if invoice_kind in {"ocr", "pdf-ocr"}:
            invoice_kind = f"{invoice_kind}+vision-fallback"

    candidates = extract_invoice_rows_from_text(invoice_text or "")
    parsed = parse_invoice_text(invoice_text or "")
    for row in parsed:
        row["archivo_origen"] = name
        row.setdefault("extraction_source", "regex")

    ollama_message = None
    ollama_tax_included = None
    settings = ollama_settings or {}
    mode = str(settings.get("mode") or "fallback")
    if should_call_ollama(
        bool(settings.get("enabled")),
        mode,
        parsed,
        invoice_text or "",
    ):
        digest = hashlib.sha256(uploaded_file.getvalue()).hexdigest()[:20]
        cache_key = "|".join(
            [
                name,
                digest,
                str(settings.get("host") or ""),
                str(settings.get("model") or ""),
                mode,
            ]
        )
        cache = st.session_state.setdefault("ollama_extract_cache", {})
        if cache_key in cache:
            ollama_result = cache[cache_key]
        else:
            ollama_result = extract_invoice_items(
                invoice_text or "",
                host=settings.get("host"),
                model=settings.get("model"),
                timeout=float(settings.get("timeout") or 120),
                source_name=name,
            )
            cache[cache_key] = ollama_result
        ollama_message = ollama_result.get("message")
        ollama_tax_included = ollama_result.get("tax_included")
        parsed = select_extracted_rows(
            parsed,
            ollama_result.get("rows") or [],
            mode=mode,
        )

    return {
        "name": name,
        "kind": invoice_kind,
        "text": invoice_text or "",
        "rows": parsed,
        "candidates": candidates,
        "columns": None,
        "ollama_message": ollama_message,
        "ollama_tax_included": ollama_tax_included,
        "error": vision_error or None,
    }


def create_log_download(logger_path):
    return Path(logger_path).read_bytes()


def render_download_panel(auto_open_folder: bool = False, key_prefix: str = "main"):
    """Panel de descarga nativo del navegador (st.download_button)."""
    if not st.session_state.get("ready") or not st.session_state.get("xlsx_bytes"):
        return

    xlsx_bytes = st.session_state["xlsx_bytes"]
    xlsx_name = st.session_state.get("xlsx_name", "Effi_Importacion.xlsx")
    size_mb = len(xlsx_bytes) / (1024 * 1024)

    st.markdown("### Descargar archivo para Effi")
    st.caption(f"{xlsx_name} · {size_mb:.2f} MB (límite Effi: 5 MB)")

    st.download_button(
        label="Descargar Excel ahora (.xlsx)",
        data=xlsx_bytes,
        file_name=xlsx_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_effi_xlsx_{key_prefix}",
        type="primary",
        use_container_width=True,
    )
    if st.session_state.get("log_bytes"):
        st.download_button(
            label="Descargar log (.log)",
            data=st.session_state["log_bytes"],
            file_name=st.session_state.get("log_name", "effi.log"),
            mime="text/plain",
            key=f"dl_effi_log_{key_prefix}",
            use_container_width=True,
        )

    path = st.session_state.get("xlsx_path")
    if path:
        st.info(f"Copia local: `{path}`")
        if auto_open_folder and os.name == "nt":
            try:
                os.startfile(str(Path(path).parent))  # noqa: S606
            except Exception:
                pass


def build_audit(
    main_rows, audit_records, omitted, conversions, bonuses, discounts, charges, summary
):
    audit = {
        "Resumen": pd.DataFrame([summary]),
        "Auditoria_Items": pd.DataFrame(audit_records),
        "Items no Encontrados": pd.DataFrame(omitted),
        "Conversiones_Unidades": pd.DataFrame(conversions),
        "Bonificaciones": pd.DataFrame(bonuses),
        "Descuentos_Individuales": pd.DataFrame(discounts),
        "Cargos_Retenciones": pd.DataFrame(charges),
    }
    # Guarantee required secondary sheet exists even if empty.
    for k in list(audit):
        if audit[k].empty:
            audit[k] = pd.DataFrame([{"Estado": "Sin registros detectados"}])
    return audit


def configure_logging(log_path):
    logger = logging.getLogger("effi_processor")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    return logger


def process(catalog_df, catalog_cols, invoice_df, settings, logger):
    tax_rate = settings["tax_rate"]
    tax_mode = settings["tax_mode"]
    match_threshold = settings["match_threshold"]

    main_rows = []
    audit_records = []
    omitted = []
    conversions = []
    bonuses = []
    discounts = []
    charges = []
    total_invoice = 0.0
    total_exported = 0.0
    total_omitted = 0.0

    catalog_index = build_catalog_index(catalog_df, catalog_cols)
    logger.info(
        "Índice de catálogo listo: %s GTINs, %s textos para fuzzy.",
        len(catalog_index["gtin_map"]),
        len(catalog_index["texts"]),
    )

    for i, raw in invoice_df.iterrows():
        row = raw.to_dict()
        desc = str(row.get("description", "") or "")
        pres = str(row.get("presentation", "") or "")
        code = clean_product_code(row.get("code", ""))

        # Exact GTIN first (mapa O(1), sin filtrar el DataFrame en cada línea)
        catalog_row = None
        score = 0
        method = ""
        if code and code in catalog_index["gtin_map"]:
            catalog_row = catalog_df.loc[catalog_index["gtin_map"][code]]
            score = 1.0
            method = "GTIN exacto"

        if catalog_row is None and code in catalog_index.get("reference_map", {}):
            catalog_row = catalog_df.loc[catalog_index["reference_map"][code]]
            score = 1.0
            method = "Código exacto"

        if catalog_row is None:
            catalog_row, score, method = match_catalog_item(
                desc, pres, catalog_df, catalog_cols, catalog_index=catalog_index
            )

        qty = parse_number(row.get("quantity")) or 0.0
        bonus_qty = parse_number(
            row.get("bonus_quantity", row.get("cantidad_bonificada"))
        ) or 0.0
        unit_inv = parse_number(row.get("unit_price_invoice")) or 0.0
        total = parse_number(row.get("total_invoice"))
        if total is None:
            total = qty * unit_inv

        # Detect pure gift lines.
        line_text = normalize_text(row.get("source_line", ""))
        is_bonus = (
            bool(row.get("is_bonus"))
            or "*" in str(row.get("source_line", ""))
            or any(
                x in line_text for x in ["bonificacion", "obsequio", "gratis", "regalo"]
            )
        )

        if catalog_row is None or score < match_threshold:
            total_omitted += total or 0
            omitted.append(
                {
                    "Línea documento": i + 1,
                    "Código documento": code,
                    "Descripción": desc,
                    "Presentación": pres,
                    "Cantidad": qty,
                    "Precio unitario documento": unit_inv,
                    "Total documento": total,
                    "Motivo": f"No encontrado / confianza {score:.1%}",
                    "Acción requerida": "CREAR PREVIAMENTE EN EFFI y luego reprocesar/importar",
                }
            )
            audit_records.append(
                {
                    "Línea": i + 1,
                    "Estado": "OMITIDO",
                    "Descripción": desc,
                    "GTIN": code,
                    "Confianza": score,
                    "Método": method,
                    "Cantidad": qty,
                    "Total factura": total,
                }
            )
            logger.warning("Item omitido línea %s: %s | score=%.3f", i + 1, desc, score)

            # Mantener la línea en el Excel final como registro editable aunque no haya coincidencia
            # exacta con el catálogo. Esto evita que el documento salga vacío y permite corregirlo.
            main_rows.append([
                code or "",
                "",
                "",
                "",
                desc,
                round(float(qty or 0), 6),
                round(float(unit_inv or 0), 6),
                0.0,
                "",
            ])
            continue

        cdesc = str(catalog_row.get(catalog_cols["description"], desc))
        cpres = str(catalog_row.get(catalog_cols["presentation"], "") or "")
        gtin = (
            clean_gtin(catalog_row.get(catalog_cols["gtin"], ""))
            if catalog_cols["gtin"]
            else ""
        )
        reference = (
            clean_product_code(catalog_row.get(catalog_cols["reference"], ""))
            if catalog_cols.get("reference")
            else ""
        )
        effi_id = (
            str(catalog_row.get(catalog_cols["effi"], "") or "")
            if catalog_cols["effi"]
            else ""
        )

        # Presentation conversion. Convert the document presentation to the catalog base presentation.
        doc_base, doc_kind = presentation_base_quantity(f"{pres} {desc}")
        cat_base, cat_kind = presentation_base_quantity(f"{cpres} {cdesc}")
        qty_adjusted = qty
        bonus_qty_adjusted = bonus_qty
        unit_conversion_factor = 1.0

        if doc_base and cat_base and doc_kind == cat_kind and cat_base > 0:
            # A line quantity is transformed into equivalent catalog presentations.
            unit_conversion_factor = doc_base / cat_base
            qty_adjusted = qty * unit_conversion_factor
            bonus_qty_adjusted = bonus_qty * unit_conversion_factor
            if abs(unit_conversion_factor - 1.0) > 1e-9:
                # Preserve total paid: unit cost becomes proportional to the adjusted physical quantity.
                conversions.append(
                    {
                        "Línea": i + 1,
                        "Documento": f"{pres} {desc}".strip(),
                        "Maestro": f"{cpres} {cdesc}".strip(),
                        "Tipo": doc_kind,
                        "Factor": unit_conversion_factor,
                        "Cantidad documento": qty,
                        "Cantidad equivalente maestro": qty_adjusted,
                        "Regla": "Equivalencia matemática; total pagado conservado",
                    }
                )
                logger.info(
                    "Conversión línea %s: factor %.9f", i + 1, unit_conversion_factor
                )

        # Determine tax base.
        if tax_mode == "Incluye IVA":
            base_total = total / (1 + tax_rate)
        elif tax_mode == "Neto":
            base_total = total
        else:  # Auto
            base_total = (
                total / (1 + tax_rate) if settings["auto_tax_included"] else total
            )

        # Discount handling: preserve item-specific discount and never apply a global discount.
        disc_pct = parse_number(row.get("discount_pct")) or 0.0
        disc_value = parse_number(row.get("discount_value")) or 0.0

        # If only percentage exists, calculate from gross invoice line.
        if disc_value == 0 and disc_pct:
            disc_value = total * disc_pct / 100.0

        # Bonus-only item: 100% discount. If it is a free unit attached to a paid reference,
        # it should preferably be merged before import; otherwise it is kept as an audit record.
        if is_bonus:
            disc_value = abs(base_total)
            disc_pct = 100.0
            bonuses.append(
                {
                    "Línea": i + 1,
                    "Descripción": desc,
                    "Cantidad obsequiada": qty,
                    "Costo asignado": base_total,
                    "Descuento": disc_value,
                    "Tratamiento": "100% descuento por bonificación pura",
                }
            )
            # Pure bonus with no direct cost: retain as a line only if catalog exists.
            # The preferred grouped X+Y treatment is handled below when explicit quantities are known.

        physical_qty = qty_adjusted + bonus_qty_adjusted
        if physical_qty > 0:
            unit_base = base_total / physical_qty
        else:
            unit_base = 0.0

        if disc_value == 0 and disc_pct == 0:
            disc_value = 0.0

        # Código Effi Impuesto: desde maestro Effi (código explícito o tarifa precio/precio+impuesto).
        tax_code, tax_source, inferred_rate = resolve_effi_tax_code(
            catalog_row, catalog_cols, settings
        )

        main_rows.append(
            [
                effi_id or gtin or reference or code,
                "",
                "",
                "",
                cdesc,
                round(physical_qty, 6),
                round(unit_base, 6),
                round(disc_value, 6),
                tax_code,
            ]
        )

        total_invoice += total
        total_exported += total
        discounts.append(
            {
                "Línea": i + 1,
                "Descripción": cdesc,
                "Porcentaje descuento": disc_pct,
                "Valor descuento": disc_value,
                "Fuente": "Documento / cálculo por línea",
            }
        )
        audit_records.append(
            {
                "Línea": i + 1,
                "Estado": "EXPORTADO",
                "Descripción": cdesc,
                "GTIN": gtin,
                "ID EFFI": effi_id,
                "Confianza": score,
                "Método": method,
                "Cantidad original": qty,
                "Cantidad ajustada": qty_adjusted,
                "Cantidad bonificada": bonus_qty_adjusted,
                "Cantidad física total": physical_qty,
                "Precio documento": unit_inv,
                "Total documento": total,
                "Base neta": base_total,
                "Descuento": disc_value,
                "Precio base unitario": unit_base,
                "IVA": inferred_rate if inferred_rate is not None else tax_rate,
                "Código impuesto": tax_code,
                "Origen impuesto": tax_source,
            }
        )

    charges.extend(settings.get("additional_charges", []))

    # The exported subtotal intentionally does not force a global balancing adjustment.
    # This is critical when omitted items exist.
    deviation = total_invoice - total_exported
    summary = {
        "Total detectado documento": round(total_invoice, 2),
        "Total representado por líneas exportadas": round(total_exported, 2),
        "Total de líneas omitidas": round(total_omitted, 2),
        "Diferencia explicada por omisiones": round(deviation, 2),
        "IVA usado": tax_rate,
        "Modo IVA": tax_mode,
        "Umbral de coincidencia": match_threshold,
        "Filas exportadas": len(main_rows),
        "Filas omitidas": len(omitted),
        "Resultado": (
            "REVISAR: existen omisiones o diferencias"
            if abs(deviation) > 0.01
            else "CUADRATURA ACEPTADA"
        ),
    }

    return (
        main_rows,
        build_audit(
            main_rows,
            audit_records,
            omitted,
            conversions,
            bonuses,
            discounts,
            charges,
            summary,
        ),
        summary,
    )


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption(
    "Procesamiento conservador: la hoja Plantilla_Importacion nunca contiene fórmulas; "
    "toda auditoría se realiza en hojas secundarias y logs. "
    "Exportación Effi: xlsx|xlsm|xls|xlt · máximo 5 MB."
)

# Descargas siempre arriba cuando hay resultado
render_download_panel(
    auto_open_folder=bool(st.session_state.pop("open_salidas", False)),
    key_prefix="top",
)

ocr_ok, ocr_msg = configure_tesseract()
ollama_ok, ollama_msg = check_ollama_vision()

with st.sidebar:
    st.header("Configuración")
    read_mode = st.selectbox(
        "Lectura de documentos",
        [
            "Auto (Vision si hay Ollama)",
            "Solo Vision",
            "Solo Tesseract",
        ],
        help=(
            "Vision usa Ollama en local (gratis, sin API cloud). "
            f"Modelo: {ollama_vision_model()}."
        ),
    )
    tax_mode = st.selectbox("Tratamiento del IVA", ["Auto", "Incluye IVA", "Neto"])
    tax_rate = st.number_input(
        "IVA",
        min_value=0.0,
        max_value=1.0,
        value=0.19,
        step=0.01,
        format="%.2f",
    )
    default_tax_code = st.text_input(
        "Código Effi Impuesto por defecto (si no se puede inferir)",
        value="1",
        help="En tus plantillas Effi, IVA 19% usa código 1. El maestro puede sobrescribirlo.",
    ).strip() or "1"
    auto_tax_included = st.checkbox(
        "En modo Auto, asumir IVA incluido si el documento lo indica", True
    )
    match_threshold = st.slider("Umbral mínimo de coincidencia", 0.50, 0.99, 0.82, 0.01)
    st.info("Los productos bajo el umbral NO se agregan a la hoja principal.")
    st.caption("Importación Effi: formatos xlsx · xlsm · xls · xlt · máx. 5 MB")
    st.caption(
        "Privacidad Vision: las facturas se procesan en localhost (Ollama); "
        "no se envían a APIs en la nube."
    )
    if ollama_ok:
        st.success(f"Vision: {ollama_msg}")
    else:
        st.warning(f"Vision: {ollama_msg}")
        st.code(f"ollama pull {ollama_vision_model()}", language="bash")
    if ocr_ok:
        st.success(f"OCR: {ocr_msg}")
    else:
        st.warning(f"OCR no disponible: {ocr_msg}")
        st.caption("Puede seguir usando Excel/CSV/TXT o PDF con texto seleccionable.")

    st.divider()
    st.subheader("Ollama (local, gratuito)")
    st.caption(
        "Usa un modelo en su PC. No envía facturas a APIs de pago. "
        "Instale https://ollama.com y ejecute `ollama pull llama3.2`."
    )
    ollama_host = st.text_input(
        "Servidor Ollama",
        value=default_host(),
        help="Por defecto http://127.0.0.1:11434",
        key="ollama_host",
    )

    ollama_status = cached_check_ollama(ollama_host)
    if ollama_status["ok"]:
        st.success(ollama_status["message"])
    else:
        st.warning(ollama_status["message"])

    available_models = ollama_status.get("models") or []
    preferred = default_model()
    if available_models:
        model_index = 0
        for i, name in enumerate(available_models):
            if name == preferred or name.startswith(preferred + ":"):
                model_index = i
                break
        ollama_model = st.selectbox(
            "Modelo",
            available_models,
            index=model_index,
            key="ollama_model",
        )
    else:
        ollama_model = st.text_input("Modelo", value=preferred, key="ollama_model_text")

    use_ollama = st.checkbox(
        "Usar Ollama para extraer líneas de PDF/imagen/TXT",
        value=bool(ollama_status["ok"]),
        disabled=not ollama_status["ok"],
        key="use_ollama",
    )
    ollama_mode_label = st.radio(
        "Cuándo usarlo",
        [
            "Respaldo (solo si el parser no obtiene líneas)",
            "Siempre (priorizar Ollama)",
        ],
        index=1,
        disabled=not use_ollama,
        key="ollama_mode",
    )
    ollama_mode = "always" if ollama_mode_label.startswith("Siempre") else "fallback"
    ollama_settings = {
        "enabled": bool(use_ollama and ollama_status["ok"]),
        "host": ollama_host,
        "model": ollama_model,
        "mode": ollama_mode,
        "timeout": 120,
    }

catalog_file = st.file_uploader(
    "1) Catálogo maestro Effi (CSV/XLSX)",
    type=["csv", "xlsx", "xls", "xlsm", "xlt"],
)
invoice_files = st.file_uploader(
    "2) Facturas / listados (puede subir varios a la vez)",
    type=INVOICE_UPLOAD_TYPES,
    accept_multiple_files=True,
    help="PDF, imagen, CSV o Excel. Varios archivos se consolidan en una sola tabla.",
)

if catalog_file:
    try:
        catalog = read_catalog(catalog_file)
        ccols = detect_catalog_columns(catalog)
        st.success(f"Catálogo cargado: {len(catalog):,} filas.")
        st.json(ccols)
        if ccols.get("price_net") and ccols.get("price_gross"):
            st.caption(
                "Impuesto Effi: se inferirá con "
                f"`{ccols['price_net']}` vs `{ccols['price_gross']}` "
                "(ej. 19% → código 1; 0% → vacío)."
            )
        elif ccols.get("tax"):
            st.caption(f"Impuesto Effi: columna de código detectada `{ccols['tax']}`.")
        else:
            st.warning(
                "El catálogo no trae código de impuesto ni par precio/precio+impuesto. "
                f"Se usará el código por defecto `{default_tax_code}`."
            )
        missing = [k for k, v in ccols.items() if v is None and k in ("gtin", "description")]
        # description o nombre es esencial; gtin puede faltar si hay ID Effi
        if ccols.get("description") is None:
            st.error("Columna de descripción/nombre no detectada. Revise el catálogo.")
        if ccols.get("gtin") is None and ccols.get("effi") is None:
            st.error("No se detectó GTIN ni ID Effi. Revise el catálogo.")
    except Exception as e:
        st.error(f"Error leyendo catálogo: {e}")
        catalog = None
        ccols = None
else:
    catalog = None
    ccols = None

invoice_df = None
invoice_text = ""
invoice_kinds = []
loaded_files = []

if invoice_files:
    all_rows = []
    all_candidates = []
    errors = []
    for uploaded in invoice_files:
        try:
            with st.spinner(f"Cargando y leyendo {uploaded.name}..."):
                meta = load_invoice_uploaded_file(
                    uploaded,
                    ollama_settings=ollama_settings,
                    read_mode=read_mode,
                )
            loaded_files.append(meta)
            invoice_kinds.append(meta["kind"])
            if meta["text"]:
                invoice_text += f"\n\n===== {meta['name']} =====\n{meta['text']}"
            all_rows.extend(meta["rows"])
            for cand in meta["candidates"]:
                cand = dict(cand)
                cand["archivo_origen"] = meta["name"]
                all_candidates.append(cand)
            st.success(
                f"{meta['name']}: {len(meta['rows'])} línea(s) · modo {meta['kind']}"
            )
            if meta.get("ollama_message"):
                st.caption(f"Ollama — {meta['name']}: {meta['ollama_message']}")
            if meta["kind"] != "spreadsheet" and meta["text"]:
                with st.expander(f"Texto detectado — {meta['name']}"):
                    st.text(meta["text"][:15000])
            if meta["columns"]:
                with st.expander(f"Columnas detectadas — {meta['name']}"):
                    st.json(meta["columns"])
        except Exception as e:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = Path(f"effi_log_error_{timestamp}.log")
            logger = configure_logging(log_path)
            logger.exception("Error leyendo factura '%s': %s", uploaded.name, e)
            errors.append((uploaded.name, str(e), log_path))
            st.error(f"Error leyendo {uploaded.name}: {e}")
            st.download_button(
                f"Descargar log — {uploaded.name}",
                data=log_path.read_bytes(),
                file_name=log_path.name,
                mime="text/plain",
                key=f"err_log_{uploaded.name}_{timestamp}",
            )

    invoice_df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame()
    if invoice_df.empty:
        st.warning(
            "No se pudo estructurar automáticamente ningún documento. "
            "Revise el texto detectado o suba las facturas en XLSX/CSV "
            "(formato Effi: xlsx|xlsm|xls|xlt · máx. 5 MB)."
        )
        if all_candidates:
            st.caption(
                f"Líneas candidatas detectadas (sin parseo completo): {len(all_candidates)}"
            )
            st.dataframe(
                pd.DataFrame(all_candidates), use_container_width=True, height=240
            )
    else:
        st.info(
            f"Consolidado: **{len(invoice_df)}** líneas de **{len(loaded_files)}** archivo(s)."
        )

if catalog is not None and invoice_df is not None and not invoice_df.empty:
    st.subheader("Previsualización consolidada (todas las facturas)")
    display_cols = [
        c
        for c in [
            "archivo_origen",
            "extraction_source",
            "code",
            "description",
            "presentation",
            "quantity",
            "bonus_quantity",
            "unit_price_invoice",
            "total_invoice",
            "discount_pct",
            "is_bonus",
            "source_line",
        ]
        if c in invoice_df.columns
    ]

    display_labels = {
        "archivo_origen": "Archivo origen",
        "code": "Código",
        "description": "Nombre / descripción",
        "presentation": "Presentación",
        "quantity": "Cantidad",
        "bonus_quantity": "Cantidad bonificada",
        "unit_price_invoice": "Valor unitario",
        "total_invoice": "Valor total",
        "discount_pct": "Descuento %",
        "is_bonus": "Es bonificación",
        "source_line": "Línea fuente",
    }

    display_df = invoice_df[display_cols].rename(columns=display_labels)
    edited_display_df = st.data_editor(
        display_df,
        use_container_width=True,
        height=360,
        num_rows="dynamic",
        key="invoice_editor",
        column_config={
            "Archivo origen": st.column_config.TextColumn("Archivo origen"),
            "Código": st.column_config.TextColumn("Código"),
            "Nombre / descripción": st.column_config.TextColumn("Nombre / descripción"),
            "Presentación": st.column_config.TextColumn("Presentación"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=0, step=1),
            "Cantidad bonificada": st.column_config.NumberColumn(
                "Cantidad bonificada", min_value=0, step=1
            ),
            "Valor unitario": st.column_config.NumberColumn("Valor unitario", format="%.2f"),
            "Valor total": st.column_config.NumberColumn("Valor total", format="%.2f"),
            "Descuento %": st.column_config.NumberColumn("Descuento %", format="%.2f"),
            "Es bonificación": st.column_config.CheckboxColumn("Es bonificación"),
            "Línea fuente": st.column_config.TextColumn("Línea fuente"),
        },
    )
    invoice_df = edited_display_df.rename(columns={v: k for k, v in display_labels.items()})

    by_file = (
        invoice_df.groupby("archivo_origen").size().reset_index(name="líneas")
        if "archivo_origen" in invoice_df.columns
        else None
    )
    if by_file is not None and len(by_file) > 1:
        st.caption("Desglose por archivo")
        st.dataframe(by_file, use_container_width=True, hide_index=True)

    process_clicked = st.button(
        "Procesar y generar Excel para Effi",
        type="primary",
        key="btn_process_effi",
    )

    if process_clicked:
        with st.spinner("Procesando facturas y generando Excel..."):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = Path(f"effi_log_{timestamp}.log")
            logger = configure_logging(log_path)

            ollama_taxes = [
                meta.get("ollama_tax_included")
                for meta in loaded_files
                if meta.get("ollama_tax_included") is not None
            ]
            effective_auto_tax = auto_tax_included
            if tax_mode == "Auto" and ollama_taxes:
                if all(ollama_taxes):
                    effective_auto_tax = True
                elif not any(ollama_taxes):
                    effective_auto_tax = False

            settings = {
                "tax_rate": tax_rate,
                "tax_mode": tax_mode,
                "auto_tax_included": auto_tax_included,
                "default_tax_code": default_tax_code,
                "match_threshold": match_threshold,
                "additional_charges": find_additional_charges(invoice_text or ""),
            }

            try:
                names = (
                    ", ".join(f["name"] for f in loaded_files)
                    if loaded_files
                    else "(sin nombre)"
                )
                logger.info("Inicio del procesamiento.")
                logger.info("Catálogo: %s | Facturas: %s", catalog_file.name, names)
                main_rows, audit_tables, summary = process(
                    catalog, ccols, invoice_df, settings, logger
                )
                summary["Archivos procesados"] = len(loaded_files)
                summary["Nombres archivos"] = names

                wb = make_workbook(main_rows, audit_tables)
                xlsx_bytes = save_workbook(wb)
                out_name = f"Effi_Importacion_{timestamp}.xlsx"

                out_disk = PROJECT_ROOT / "salidas" / out_name
                out_disk.parent.mkdir(parents=True, exist_ok=True)
                out_disk.write_bytes(xlsx_bytes)
                log_bytes = create_log_download(log_path)
                log_disk = PROJECT_ROOT / "salidas" / log_path.name
                log_disk.write_bytes(log_bytes)

                st.session_state["ready"] = True
                st.session_state["xlsx_bytes"] = xlsx_bytes
                st.session_state["xlsx_name"] = out_name
                st.session_state["xlsx_path"] = str(out_disk)
                st.session_state["log_bytes"] = log_bytes
                st.session_state["log_name"] = log_path.name
                st.session_state["log_path"] = str(log_disk)
                st.session_state["summary"] = summary
                st.session_state["audit_tables"] = {
                    k: v.copy() for k, v in audit_tables.items()
                }
                st.session_state["process_error"] = None
                st.session_state["open_salidas"] = True
                st.success("Procesado. El botón de descarga está arriba")
                st.rerun()
            except Exception as e:
                logger.exception("Error fatal durante el procesamiento.")
                st.session_state["ready"] = False
                st.session_state["process_error"] = str(e)
                st.session_state["log_bytes"] = (
                    log_path.read_bytes() if log_path.exists() else b""
                )
                st.session_state["log_name"] = log_path.name
                st.error(f"Error durante el procesamiento: {e}")

if st.session_state.get("process_error"):
    st.error(f"Último error de procesamiento: {st.session_state['process_error']}")
    if st.session_state.get("log_bytes"):
        st.download_button(
            label="Descargar log del error",
            data=st.session_state["log_bytes"],
            file_name=st.session_state.get("log_name", "error.log"),
            mime="text/plain",
            key="dl_error_log",
        )

if st.session_state.get("ready") and st.session_state.get("summary"):
    st.divider()
    st.subheader("Resultado del procesamiento")
    summary = st.session_state.get("summary") or {}
    if summary.get("Filas omitidas", 0) > 0:
        st.error(
            f"ALERTA: {summary['Filas omitidas']} ítem(s) fueron omitidos "
            "por no alcanzar el umbral o no tener coincidencia. "
            "Deben crearse previamente en Effi."
        )
    else:
        st.success("Procesamiento completado sin omisiones.")
    st.json(summary)
    # Repetir descarga abajo por comodidad
    render_download_panel(auto_open_folder=False, key_prefix="bottom")

    audit_tables = st.session_state.get("audit_tables") or {}
    if audit_tables:
        st.subheader("Auditoría")
        for name, df in audit_tables.items():
            with st.expander(name):
                st.dataframe(df, use_container_width=True)
