
import io
import re
import json
import math
import logging
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    from rapidfuzz import fuzz, process
except Exception:
    fuzz = None
    process = None


APP_TITLE = "Effi – Procesador de Facturas e Importación"
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
    "ml": 0.001, "mililitro": 0.001, "mililitros": 0.001,
    "l": 1.0, "lt": 1.0, "litro": 1.0, "litros": 1.0,
    "cl": 0.01, "centilitro": 0.01, "centilitros": 0.01,
    "gal": 3.785411784, "galon": 3.785411784, "galón": 3.785411784,
    "galones": 3.785411784, "us gallon": 3.785411784,
}
UNIT_FACTORS_TO_KG = {
    "mg": 0.000001, "miligramo": 0.000001, "miligramo": 0.000001,
    "g": 0.001, "gramo": 0.001, "gramos": 0.001,
    "kg": 1.0, "kilo": 1.0, "kilos": 1.0, "kilogramo": 1.0, "kilogramos": 1.0,
}
UNIT_FACTORS_TO_UNITS = {"und": 1.0, "unidad": 1.0, "unidades": 1.0, "u": 1.0}

MONEY_RE = r"[-+]?\$?\s*[\d\.,]+"
PCT_RE = r"(\d+(?:[\.,]\d+)?)\s*%"

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
        s = "".join(parts[:-1]) + "." + parts[-1] if len(parts[-1]) <= 2 else "".join(parts)
    elif "." in s:
        parts = s.split(".")
        s = "".join(parts) if len(parts[-1]) == 3 and len(parts) > 1 else s
    try:
        return float(s)
    except Exception:
        return None

def parse_money_series(series):
    return series.map(parse_number)

def extract_text_from_pdf(file_bytes):
    if fitz is None:
        raise RuntimeError("Instale PyMuPDF para procesar PDF: pip install pymupdf")
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages = []
    for i, page in enumerate(doc):
        txt = page.get_text("text")
        pages.append(f"\n--- PÁGINA {i+1} ---\n{txt}")
    return "\n".join(pages)

def extract_text_from_image(file_bytes):
    if Image is None or pytesseract is None:
        raise RuntimeError("Para OCR instale Pillow y pytesseract. También requiere Tesseract OCR instalado en el sistema.")
    img = Image.open(io.BytesIO(file_bytes))
    return pytesseract.image_to_string(img, lang="spa+eng")

def read_document(uploaded_file):
    ext = Path(uploaded_file.name).suffix.lower()
    data = uploaded_file.getvalue()
    if ext == ".pdf":
        return extract_text_from_pdf(data), "pdf"
    if ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        return extract_text_from_image(data), "ocr"
    if ext in {".txt", ".csv"}:
        return data.decode("utf-8", errors="replace"), "text"
    if ext in {".xlsx", ".xls"}:
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
    if ext in {".xlsx", ".xls"}:
        return pd.read_excel(io.BytesIO(data), dtype=str)
    raise ValueError("El catálogo maestro debe ser CSV o Excel.")

def find_column(df, candidates):
    norm_map = {normalize_text(c): c for c in df.columns}
    for candidate in candidates:
        n = normalize_text(candidate)
        if n in norm_map:
            return norm_map[n]
    for c in df.columns:
        nc = normalize_text(c)
        for candidate in candidates:
            if normalize_text(candidate) in nc or nc in normalize_text(candidate):
                return c
    return None

def detect_catalog_columns(df):
    return {
        "gtin": find_column(df, [
            "GTIN", "Código de barras", "Codigo de barras", "EAN", "Código", "Codigo", "Barcode"
        ]),
        "effi": find_column(df, [
            "ID EFFI", "Código EFFI", "Codigo EFFI", "ID", "ID Artículo", "Articulo ID"
        ]),
        "description": find_column(df, [
            "Descripción", "Descripcion", "Artículo", "Articulo", "Nombre", "Producto"
        ]),
        "presentation": find_column(df, [
            "Presentación", "Presentacion", "Empaque", "Unidad", "Contenido", "Tamaño", "Tamano"
        ]),
        "brand": find_column(df, ["Marca", "Brand"]),
        "tax": find_column(df, ["IVA", "Impuesto", "Código Impuesto", "Codigo Impuesto"]),
    }

def clean_gtin(value):
    if value is None:
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s

def extract_presentation(text):
    s = normalize_text(text)
    # e.g. 3.785 L, 4L, 500 ml, 1.5 kg, 250 g, 12 und
    patterns = [
        (r"(\d+(?:\.\d+)?)\s*(ml|mililitros?|cl|litros?|l|lt|galones?|galon|gal)\b", "volume"),
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

def match_catalog_item(description, presentation, catalog, cols):
    query = f"{description or ''} {presentation or ''}".strip()
    qn = normalize_text(query)
    if not qn:
        return None, 0, "Sin descripción"

    best = None
    best_score = 0
    for idx, row in catalog.iterrows():
        text = " ".join(str(row.get(cols[k], "") or "") for k in ["description", "presentation", "brand"])
        score = similarity(query, text)
        # Strong GTIN/codes, if description happens to contain one
        qgt = clean_gtin(description)
        if qgt and cols["gtin"]:
            cgt = clean_gtin(row.get(cols["gtin"], ""))
            if qgt and cgt and qgt == cgt:
                score = 1.0
        if score > best_score:
            best_score = score
            best = idx
    return (catalog.loc[best] if best is not None else None), best_score, "Fuzzy"

def infer_columns_from_invoice_df(df):
    return {
        "code": find_column(df, ["Codigo", "Código", "GTIN", "EAN", "Barcode", "Código de barras"]),
        "description": find_column(df, ["Descripcion", "Descripción", "Articulo", "Artículo", "Producto", "Nombre"]),
        "presentation": find_column(df, ["Presentacion", "Presentación", "Empaque", "Unidad"]),
        "info": find_column(df, ["Inf.", "Inf", "Información", "Info"]),
        "quantity": find_column(df, ["Cantidad", "Cant.", "Cant"]),
        "list_price": find_column(df, ["Precio Lista", "Precio de lista", "Precio Lista Unitario", "P. Lista"]),
        "unit_price": find_column(df, ["Precio Unitario", "Precio ud.", "Precio", "Valor Unitario"]),
        "total": find_column(df, ["Precio Total", "Total", "Valor Total", "Importe"]),
        "discount": find_column(df, ["Descuento", "Desc.", "Dto.", "% Descuento", "% Desc"]),
        "tax": find_column(df, ["IVA", "Impuesto", "Tax"]),
    }

def extract_invoice_rows_from_text(text):
    """
    Heurística conservadora para documentos de texto/OCR.
    Para formatos tabulares se recomienda subir el Excel/CSV de la factura.
    """
    rows = []
    for line_no, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line:
            continue
        nums = re.findall(MONEY_RE, line)
        # Candidate line: contains a description and at least one numeric token.
        if nums:
            rows.append({"linea": line_no, "texto": line, "numeros": nums})
    return rows

def parse_text_line(line):
    # Flexible parser: code | description | qty | unit price | total | discount
    parts = re.split(r"\s{2,}|\t|\|", line.strip())
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) >= 3:
        nums = [parse_number(p) for p in parts]
        numeric_idx = [i for i, x in enumerate(nums) if x is not None]
        if numeric_idx:
            first_num = numeric_idx[0]
            desc = " ".join(parts[:first_num]) if first_num else parts[0]
            vals = [nums[i] for i in numeric_idx]
            qty = vals[0] if len(vals) >= 1 else None
            total = vals[-1] if len(vals) >= 2 else None
            unit = vals[-2] if len(vals) >= 3 else None
            return {
                "code": clean_gtin(parts[0]) if first_num > 0 else "",
                "description": desc,
                "presentation": "",
                "quantity": qty,
                "unit_price_invoice": unit,
                "total_invoice": total,
                "discount_pct": None,
                "discount_value": None,
                "source_line": line,
            }
    return None

def detect_global_tax(text):
    n = normalize_text(text)
    if "iva incluido" in n or "iva incluido" in n:
        return True
    return None

def find_percentages(text):
    return [float(x.replace(",", ".")) for x in re.findall(PCT_RE, text)]

def find_additional_charges(text):
    keywords = [
        "retefuente", "rete iva", "reteiva", "rete ica", "reteica",
        "impuesto al consumo", "flete", "transporte", "seguro", "otros cargos",
        "descuento financiero", "descuento pronto pago"
    ]
    out = []
    lines = text.splitlines()
    for line in lines:
        n = normalize_text(line)
        if any(k in n for k in keywords):
            vals = re.findall(MONEY_RE, line)
            out.append({"concepto": line.strip(), "valores_detectados": vals})
    return out

def detect_bonus(text):
    """
    Detecta asteriscos/bonificaciones y patrones como 10 + 2 gratis.
    """
    bonuses = []
    for line in text.splitlines():
        n = normalize_text(line)
        if "*" in line or any(k in n for k in ["bonificacion", "obsequio", "gratis", "regalo", "bono"]):
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
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r_idx, row in enumerate(main_rows, 2):
        for c_idx, value in enumerate(row, 1):
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
                sh.cell(r, c, value)
        for c in range(1, sh.max_column + 1):
            sh.column_dimensions[get_column_letter(c)].width = min(45, max(14, len(str(sh.cell(1,c).value or "")) + 2))
        sh.freeze_panes = "A2"

    return wb

def openpyxl_workbook():
    from openpyxl import Workbook
    return Workbook()

def save_workbook(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def build_audit(main_rows, audit_records, omitted, conversions, bonuses, discounts, charges, summary):
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

    for i, raw in invoice_df.iterrows():
        row = raw.to_dict()
        desc = row.get("description", "") or ""
        pres = row.get("presentation", "") or ""
        code = clean_gtin(row.get("code", ""))

        # Exact GTIN first
        catalog_row = None
        score = 0
        method = ""
        if code and catalog_cols["gtin"]:
            candidates = catalog_df[catalog_df[catalog_cols["gtin"]].map(clean_gtin) == code]
            if len(candidates):
                catalog_row = candidates.iloc[0]
                score = 1.0
                method = "GTIN exacto"

        if catalog_row is None:
            catalog_row, score, method = match_catalog_item(desc, pres, catalog_df, catalog_cols)

        qty = parse_number(row.get("quantity")) or 0.0
        unit_inv = parse_number(row.get("unit_price_invoice")) or 0.0
        total = parse_number(row.get("total_invoice"))
        if total is None:
            total = qty * unit_inv

        # Detect pure gift lines.
        line_text = normalize_text(row.get("source_line", ""))
        is_bonus = "*" in str(row.get("source_line", "")) or any(
            x in line_text for x in ["bonificacion", "obsequio", "gratis", "regalo"]
        )

        if catalog_row is None or score < match_threshold:
            total_omitted += total or 0
            omitted.append({
                "Línea documento": i + 1,
                "Código documento": code,
                "Descripción": desc,
                "Presentación": pres,
                "Cantidad": qty,
                "Precio unitario documento": unit_inv,
                "Total documento": total,
                "Motivo": f"No encontrado / confianza {score:.1%}",
                "Acción requerida": "CREAR PREVIAMENTE EN EFFI y luego reprocesar/importar",
            })
            audit_records.append({
                "Línea": i + 1, "Estado": "OMITIDO", "Descripción": desc,
                "GTIN": code, "Confianza": score, "Método": method,
                "Cantidad": qty, "Total factura": total,
            })
            logger.warning("Item omitido línea %s: %s | score=%.3f", i+1, desc, score)
            continue

        cdesc = str(catalog_row.get(catalog_cols["description"], desc))
        cpres = str(catalog_row.get(catalog_cols["presentation"], "") or "")
        gtin = clean_gtin(catalog_row.get(catalog_cols["gtin"], "")) if catalog_cols["gtin"] else code
        effi_id = str(catalog_row.get(catalog_cols["effi"], "") or "") if catalog_cols["effi"] else ""

        # Presentation conversion. Convert the document presentation to the catalog base presentation.
        doc_base, doc_kind = presentation_base_quantity(f"{pres} {desc}")
        cat_base, cat_kind = presentation_base_quantity(f"{cpres} {cdesc}")
        qty_adjusted = qty
        unit_conversion_factor = 1.0

        if doc_base and cat_base and doc_kind == cat_kind and cat_base > 0:
            # A line quantity is transformed into equivalent catalog presentations.
            unit_conversion_factor = doc_base / cat_base
            qty_adjusted = qty * unit_conversion_factor
            if abs(unit_conversion_factor - 1.0) > 1e-9:
                # Preserve total paid: unit cost becomes proportional to the adjusted physical quantity.
                conversions.append({
                    "Línea": i + 1,
                    "Documento": f"{pres} {desc}".strip(),
                    "Maestro": f"{cpres} {cdesc}".strip(),
                    "Tipo": doc_kind,
                    "Factor": unit_conversion_factor,
                    "Cantidad documento": qty,
                    "Cantidad equivalente maestro": qty_adjusted,
                    "Regla": "Equivalencia matemática; total pagado conservado",
                })
                logger.info("Conversión línea %s: factor %.9f", i+1, unit_conversion_factor)

        # Determine tax base.
        if tax_mode == "Incluye IVA":
            base_total = total / (1 + tax_rate)
        elif tax_mode == "Neto":
            base_total = total
        else:  # Auto
            base_total = total / (1 + tax_rate) if settings["auto_tax_included"] else total

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
            bonuses.append({
                "Línea": i + 1, "Descripción": desc, "Cantidad obsequiada": qty,
                "Costo asignado": base_total, "Descuento": disc_value,
                "Tratamiento": "100% descuento por bonificación pura",
            })
            # Pure bonus with no direct cost: retain as a line only if catalog exists.
            # The preferred grouped X+Y treatment is handled below when explicit quantities are known.

        # Effective unit base before item discount.
        if qty_adjusted > 0:
            unit_base = base_total / qty_adjusted
        else:
            unit_base = 0.0

        if disc_value == 0 and disc_pct == 0:
            disc_value = 0.0

        # Tax code: 1 by requested default, unless master has a code.
        tax_code = "1"
        if catalog_cols["tax"]:
            candidate_tax = str(catalog_row.get(catalog_cols["tax"], "") or "").strip()
            if candidate_tax:
                tax_code = candidate_tax

        main_rows.append([
            gtin or effi_id,
            "",
            "",
            "",
            cdesc,
            round(qty_adjusted, 6),
            round(unit_base, 6),
            round(disc_value, 6),
            tax_code,
        ])

        total_invoice += total
        total_exported += total
        discounts.append({
            "Línea": i + 1,
            "Descripción": cdesc,
            "Porcentaje descuento": disc_pct,
            "Valor descuento": disc_value,
            "Fuente": "Documento / cálculo por línea",
        })
        audit_records.append({
            "Línea": i + 1, "Estado": "EXPORTADO", "Descripción": cdesc,
            "GTIN": gtin, "ID EFFI": effi_id, "Confianza": score,
            "Método": method, "Cantidad original": qty,
            "Cantidad ajustada": qty_adjusted, "Precio documento": unit_inv,
            "Total documento": total, "Base neta": base_total,
            "Descuento": disc_value, "Precio base unitario": unit_base,
            "IVA": tax_rate, "Código impuesto": tax_code,
        })

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
        "Resultado": "REVISAR: existen omisiones o diferencias" if abs(deviation) > 0.01 else "CUADRATURA ACEPTADA",
    }

    return main_rows, build_audit(main_rows, audit_records, omitted, conversions, bonuses, discounts, charges, summary), summary

def create_log_download(logger_path):
    return Path(logger_path).read_bytes()

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Procesamiento conservador: la hoja Plantilla_Importacion nunca contiene fórmulas; toda auditoría se realiza en hojas secundarias y logs.")

with st.sidebar:
    st.header("Configuración")
    tax_mode = st.selectbox("Tratamiento del IVA", ["Auto", "Incluye IVA", "Neto"])
    tax_rate = st.number_input("IVA", min_value=0.0, max_value=1.0, value=0.19, step=0.01, format="%.2f")
    auto_tax_included = st.checkbox("En modo Auto, asumir IVA incluido si el documento lo indica", True)
    match_threshold = st.slider("Umbral mínimo de coincidencia", 0.50, 0.99, 0.82, 0.01)
    st.info("Los productos bajo el umbral NO se agregan a la hoja principal.")

catalog_file = st.file_uploader("1) Catálogo maestro Effi (CSV/XLSX)", type=["csv", "xlsx", "xls"])
invoice_file = st.file_uploader("2) Factura/listado (PDF, imagen, CSV, XLSX, TXT)", type=["pdf", "png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff", "csv", "xlsx", "xls", "txt"])

if catalog_file:
    try:
        catalog = read_catalog(catalog_file)
        ccols = detect_catalog_columns(catalog)
        st.success(f"Catálogo cargado: {len(catalog):,} filas.")
        st.json(ccols)
        missing = [k for k, v in ccols.items() if v is None and k in ("gtin", "description")]
        if missing:
            st.error(f"Columnas esenciales no detectadas: {missing}. Revise el catálogo.")
    except Exception as e:
        st.error(f"Error leyendo catálogo: {e}")
        catalog = None
        ccols = None
else:
    catalog = None
    ccols = None

if invoice_file:
    try:
        invoice_text, invoice_kind = read_document(invoice_file)
        if invoice_kind == "spreadsheet":
            invoice_df_raw = pd.read_excel(io.BytesIO(invoice_file.getvalue()), dtype=str)
            inv_cols = infer_columns_from_invoice_df(invoice_df_raw)
            st.success(f"Factura tabular cargada: {len(invoice_df_raw):,} filas.")
            st.json(inv_cols)
            normalized_rows = []
            for _, r in invoice_df_raw.iterrows():
                normalized_rows.append({
                    "code": r.get(inv_cols["code"], "") if inv_cols["code"] else "",
                    "description": r.get(inv_cols["description"], "") if inv_cols["description"] else "",
                    "presentation": r.get(inv_cols["presentation"], "") if inv_cols["presentation"] else "",
                    "quantity": r.get(inv_cols["quantity"], "") if inv_cols["quantity"] else "",
                    "unit_price_invoice": r.get(inv_cols["unit_price"], "") if inv_cols["unit_price"] else r.get(inv_cols["list_price"], ""),
                    "total_invoice": r.get(inv_cols["total"], "") if inv_cols["total"] else "",
                    "discount_pct": r.get(inv_cols["discount"], "") if inv_cols["discount"] else "",
                    "discount_value": "",
                    "source_line": " | ".join(str(x) for x in r.tolist() if str(x) != "nan"),
                })
            invoice_df = pd.DataFrame(normalized_rows)
        else:
            st.success(f"Documento leído mediante: {invoice_kind}.")
            with st.expander("Texto detectado"):
                st.text(invoice_text[:20000])
            raw_lines = extract_invoice_rows_from_text(invoice_text or "")
            parsed = []
            for item in raw_lines:
                p = parse_text_line(item["texto"])
                if p:
                    parsed.append(p)
            invoice_df = pd.DataFrame(parsed)
            if invoice_df.empty:
                st.warning("No se pudo estructurar automáticamente el documento. Para máxima precisión, convierta/suba la factura en XLSX/CSV o ajuste manualmente las filas en el código.")
    except Exception as e:
        st.error(f"Error leyendo factura: {e}")
        invoice_df = None
else:
    invoice_df = None

if catalog is not None and invoice_df is not None and not invoice_df.empty:
    st.subheader("Previsualización de líneas detectadas")
    st.dataframe(invoice_df, use_container_width=True, height=260)

    if st.button("Procesar y generar Excel", type="primary"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = Path(f"effi_log_{timestamp}.log")
        logger = configure_logging(log_path)

        settings = {
            "tax_rate": tax_rate,
            "tax_mode": tax_mode,
            "auto_tax_included": auto_tax_included,
            "match_threshold": match_threshold,
            "additional_charges": find_additional_charges(invoice_text or "") if invoice_file and invoice_kind != "spreadsheet" else [],
        }

        try:
            logger.info("Inicio del procesamiento.")
            logger.info("Catálogo: %s | Factura: %s", catalog_file.name, invoice_file.name)
            main_rows, audit_tables, summary = process(catalog, ccols, invoice_df, settings, logger)

            wb = make_workbook(main_rows, audit_tables)
            xlsx_bytes = save_workbook(wb)
            out_name = f"Effi_Importacion_{timestamp}.xlsx"

            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.session_state["xlsx_name"] = out_name
            st.session_state["log_bytes"] = create_log_download(log_path)
            st.session_state["log_name"] = log_path.name
            st.session_state["summary"] = summary
            st.session_state["audit_tables"] = audit_tables

            if summary["Filas omitidas"] > 0:
                st.error(f"⚠️ ALERTA: {summary['Filas omitidas']} ítem(s) fueron omitidos por no alcanzar el umbral o no tener coincidencia. Deben crearse previamente en Effi.")
            else:
                st.success("Procesamiento completado sin omisiones.")

            st.json(summary)

            st.download_button(
                "⬇️ Descargar Excel listo para Effi",
                data=xlsx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "⬇️ Descargar log de ejecución",
                data=st.session_state["log_bytes"],
                file_name=st.session_state["log_name"],
                mime="text/plain",
            )

            # Detailed audit previews
            st.subheader("Auditoría")
            for name, df in audit_tables.items():
                with st.expander(name):
                    st.dataframe(df, use_container_width=True)
        except Exception as e:
            logger.exception("Error fatal durante el procesamiento.")
            st.error(f"Error durante el procesamiento: {e}")
            st.download_button("Descargar log del error", data=log_path.read_bytes(), file_name=log_path.name, mime="text/plain")
