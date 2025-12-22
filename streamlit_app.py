import io
import re
import zipfile
from datetime import datetime
from typing import List, Dict, Optional, Any, Tuple
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st


# =========================================================
# SECURITY / PRIVACY
# =========================================================
# - No uploads are written to disk
# - No caching of uploaded data
# - Outputs are generated in memory only
# =========================================================


# =========================================================
# SYSTEM TYPES
# =========================================================
SYSTEM_TYPES = ["Cadre", "Voelkr"]


# =========================================================
# CADRE OUTPUT COLUMNS (your existing)
# =========================================================
CADRE_COLUMNS = [
    "ReferralManager",
    "ReferralEmail",
    "Brand",
    "QuoteNumber",
    "QuoteDate",
    "Company",
    "FirstName",
    "LastName",
    "ContactEmail",
    "ContactPhone",
    "Address",
    "County",
    "City",
    "State",
    "ZipCode",
    "Country",
    "item_id",
    "item_desc",
    "UnitPrice",
    "TotalSales",
    "QuoteValidDate",
    "CustomerNumber",
    "manufacturer_Name",
    "PDF",
    "DemoQuote",
]


# =========================================================
# VOELKR OUTPUT COLUMNS (PUT YOUR UPLOADED CSV HEADERS HERE)
# =========================================================
# Replace the list below with the EXACT header columns from your Voelkr CSV.
# Example placeholders shown:
VOELKR_FIELDS = [
    # --- paste your real CSV headers here (exact spelling/case) ---
    "PDF",
    "DocumentNumber",
    "DocumentDate",
    "CustomerName",
    "CustomerNumber",
    "ShipToAddress",
    "BillToAddress",
    "Subtotal",
    "Tax",
    "Total",
    # Add all remaining columns...
]

# =========================================================
# VOELKR FIELD EXTRACTION MAP (regex per field)
# =========================================================
# For highest accuracy, define regex per field.
# Put capturing group ( ... ) around the value you want.
# If you don't know a field's regex yet, leave it as "" and it will be blank.
VOELKR_REGEX_MAP: Dict[str, str] = {
    # Examples (you MUST tailor these to your Voelkr PDF layout):
    "DocumentNumber": r"(?:Document|Invoice|Order)\s*(?:No\.?|Number)?\s*[:#]?\s*([A-Z0-9\-]+)",
    "DocumentDate": r"(?:Date)\s*[:#]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
    "CustomerName": r"(?:Customer|Sold To|Bill To)\s*[:#]?\s*(.+)",
    "CustomerNumber": r"(?:Customer\s*No\.?|Customer\s*Number)\s*[:#]?\s*([A-Z0-9\-]+)",
    "Subtotal": r"(?:Subtotal)\s*[:#]?\s*\$?\s*([\d,]+\.\d{2})",
    "Tax": r"(?:Tax)\s*[:#]?\s*\$?\s*([\d,]+\.\d{2})",
    "Total": r"(?:Total)\s*[:#]?\s*\$?\s*([\d,]+\.\d{2})",
    "ShipToAddress": r"(?:Ship\s*To)\s*[:#]?\s*(.+)",
    "BillToAddress": r"(?:Bill\s*To)\s*[:#]?\s*(.+)",
    # Add the rest of your fields here...
}


# =========================================================
# COMMON HELPERS
# =========================================================
def normalize_date_str(date_str: Optional[str]) -> Optional[str]:
    if not date_str:
        return None
    m = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})", date_str)
    if not m:
        return date_str.strip()
    mm = int(m.group(1))
    dd = int(m.group(2))
    yy = int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def extract_full_text(pdf_bytes: bytes) -> str:
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                txt = page.extract_text(layout=True) or ""
            except TypeError:
                txt = page.extract_text() or ""
            full_text += txt + "\n"
    return full_text


# =========================================================
# CADRE LINE ITEMS (your hybrid logic – unchanged core)
# =========================================================
MONEY_RE = re.compile(r"^\$?[\d,]+\.\d{2,5}$")
QTY_RE = re.compile(r"^\d+(?:\.\d+)?$")
ITEM_ID_RE = re.compile(r"^[A-Z0-9][A-Z0-9.\-/_]+$")
ITEM_START_RE = re.compile(r"^(\d{1,4})\s+([A-Z0-9][A-Z0-9.\-/_]+)\b", re.ASCII)

SUMMARY_STOP_LINE_RE = re.compile(
    r"^(Subtotal|Total\b|Grand\s+Total|Freight|Tax\b|Product\b)",
    re.IGNORECASE,
)


def extract_header_info_cadre(full_text: str) -> Dict[str, Optional[str]]:
    header: Dict[str, Optional[str]] = {}

    m_quote = re.search(r"Quote\s+(\d+)\s+Date\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_quote:
        header["QuoteNumber"] = m_quote.group(1)
        header["QuoteDate"] = m_quote.group(2)

    m_cust = re.search(r"Customer\s+(\d+)", full_text)
    if m_cust:
        header["CustomerNumber"] = m_cust.group(1)

    m_contact = re.search(r"Contact\s+([A-Za-z .'-]+)", full_text)
    if m_contact:
        name = m_contact.group(1).strip()
        parts = name.split()
        if len(parts) >= 2:
            header["FirstName"] = parts[0]
            header["LastName"] = " ".join(parts[1:])
        elif parts:
            header["FirstName"] = parts[0]

    m_sales = re.search(r"Salesperson\s+([A-Za-z .'-]+)", full_text)
    if m_sales:
        header["ReferralManager"] = m_sales.group(1).strip()

    if "Quoted For:" in full_text and "Quote Good Through" in full_text:
        try:
            start = full_text.index("Quoted For:")
            end = full_text.index("Quote Good Through")
            addr_block = full_text[start:end]
        except Exception:
            addr_block = ""

        m_company = re.search(r"Quoted For:\s*(.+?)\s+Ship To:", addr_block)
        if m_company:
            header["Company"] = m_company.group(1).strip()

        m_addr = re.search(r"(\d{3,6}\s+[A-Za-z0-9 .#-]+)", addr_block)
        if m_addr:
            header["Address"] = m_addr.group(1).strip()

        m_city = re.search(r"([A-Za-z .]+),\s*([A-Z]{2})\s+(\d{5})(?:-\d{4})?", addr_block)
        if m_city:
            header["City"] = m_city.group(1).strip()
            header["State"] = m_city.group(2)
            header["ZipCode"] = m_city.group(3)

        if "United States of America" in addr_block:
            header["Country"] = "USA"

    m_valid = re.search(r"Quote Good Through\s+(\d{1,2}/\d{1,2}/\d{4})", full_text)
    if m_valid:
        header["QuoteValidDate"] = m_valid.group(1)

    return header


def extract_tax_amount(full_text: str) -> Optional[float]:
    block = full_text
    if "Product" in full_text and "Total" in full_text:
        try:
            start = full_text.index("Product")
            end = full_text.index("Total", start)
            block = full_text[start:end]
        except Exception:
            block = full_text

    m = re.search(r"\bTax\s+([\d,]+\.\d{2})\b", block)
    if not m:
        m = re.search(r"\bTax\s+([\d,]+\.\d{2})\b", full_text)
        if not m:
            return None
    try:
        return float(m.group(1).replace(",", ""))
    except Exception:
        return None


def _safe_float(s: Optional[str]) -> Optional[float]:
    if not s:
        return None
    try:
        return float(str(s).replace("$", "").replace(",", ""))
    except Exception:
        return None


def _clean_cell(x: Any) -> str:
    if x is None:
        return ""
    return str(x).replace("\n", " ").strip()


def _row_has_item_signature(cells: List[str]) -> bool:
    if len(cells) < 2:
        return False
    c0 = cells[0].strip()
    c1 = cells[1].strip()
    return c0.isdigit() and bool(ITEM_ID_RE.match(c1))


def _parse_item_from_table_row(cells: List[str]) -> Dict[str, str]:
    line_no = cells[0].strip()
    item_id = cells[1].strip()

    tokens: List[str] = []
    for c in cells[2:]:
        tokens.extend(c.split())

    qty = ""
    for t in tokens:
        if QTY_RE.match(t):
            qty = t
            break

    money = [t for t in tokens if MONEY_RE.match(t)]
    unit_price = money[-2] if len(money) >= 2 else (money[-1] if len(money) == 1 else "")
    total = money[-1] if len(money) >= 1 else ""

    desc_parts = []
    for c in cells[2:]:
        c_clean = c.strip()
        if not c_clean:
            continue
        toks = c_clean.split()
        non_num = [
            t for t in toks
            if not (QTY_RE.match(t) or MONEY_RE.match(t) or re.fullmatch(r"[A-Z/]+", t) or t.isdigit())
        ]
        if non_num:
            desc_parts.append(c_clean)

    return {
        "line_no": line_no,
        "item_id": item_id,
        "qty": qty,
        "unit_price": unit_price,
        "total": total,
        "description": " ".join(desc_parts).strip(),
    }


def _line_is_desc_continuation(line: str) -> bool:
    if not line:
        return False
    if ITEM_START_RE.match(line):
        return False
    if SUMMARY_STOP_LINE_RE.match(line):
        return False
    toks = line.split()
    wordy = any(
        not (QTY_RE.match(t) or MONEY_RE.match(t) or re.fullmatch(r"[A-Z/]+", t) or t.isdigit())
        for t in toks
    )
    return wordy


def extract_line_items_by_tables(pdf_bytes: bytes) -> List[Dict[str, str]]:
    items: List[Dict[str, str]] = []

    table_settings_variants = [
        {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "edge_min_length": 20,
            "min_words_vertical": 1,
            "min_words_horizontal": 1,
            "intersection_tolerance": 3,
            "text_tolerance": 2,
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "min_words_vertical": 1,
            "min_words_horizontal": 1,
            "text_tolerance": 2,
        },
    ]

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_items_found = False
            for settings in table_settings_variants:
                try:
                    tables = page.extract_tables(table_settings=settings) or []
                except Exception:
                    tables = []

                for table in tables:
                    if not table:
                        continue
                    str_rows = [[_clean_cell(c) for c in (raw_row or [])] for raw_row in table]

                    r = 0
                    while r < len(str_rows):
                        cells = str_rows[r]
                        if cells and SUMMARY_STOP_LINE_RE.match(cells[0].strip()):
                            r += 1
                            continue

                        if _row_has_item_signature(cells):
                            item = _parse_item_from_table_row(cells)

                            if not item.get("description") and r + 1 < len(str_rows):
                                nxt_cells = str_rows[r + 1]
                                nxt_line = " ".join([c for c in nxt_cells if c]).strip()
                                if _line_is_desc_continuation(nxt_line):
                                    item["description"] = nxt_line
                                    r += 1

                            items.append(item)
                            page_items_found = True

                        r += 1

                if page_items_found:
                    break

    return items


def _group_words_into_rows(words, y_tol=2.5) -> List[List[dict]]:
    buckets = defaultdict(list)
    for w in words:
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)
    return [sorted(buckets[y], key=lambda x: x["x0"]) for y in sorted(buckets.keys())]


def _row_to_text(row_words: List[dict]) -> str:
    return " ".join(w["text"] for w in row_words).strip()


def _parse_item_row_tokens(tokens: List[str]) -> Dict[str, Optional[str]]:
    line_no = tokens[0]
    item_id = tokens[1]
    after = tokens[2:]

    qty = None
    for t in after:
        if QTY_RE.match(t):
            qty = t
            break

    money_idxs = [i for i, t in enumerate(after) if MONEY_RE.match(t)]
    unit_price = None
    total = None
    if len(money_idxs) >= 2:
        total = after[money_idxs[-1]]
        unit_price = after[money_idxs[-2]]
    elif len(money_idxs) == 1:
        total = after[money_idxs[-1]]

    return {"line_no": line_no, "item_id": item_id, "qty": qty, "unit_price": unit_price, "total": total}


def extract_line_items_by_words(pdf_bytes: bytes) -> List[Dict[str, str]]:
    visual_lines: List[str] = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2, y_tolerance=2, keep_blank_chars=False, use_text_flow=True
            )
            if not words:
                continue
            for row in _group_words_into_rows(words, y_tol=2.5):
                txt = _row_to_text(row)
                if txt:
                    visual_lines.append(txt)

    items: List[Dict[str, Optional[str]]] = []
    current: Optional[Dict[str, Optional[str]]] = None
    desc_parts: List[str] = []

    def flush():
        nonlocal current, desc_parts
        if current:
            current["description"] = " ".join(desc_parts).strip()
            items.append(current)
        current = None
        desc_parts = []

    i = 0
    while i < len(visual_lines):
        line = visual_lines[i].strip()

        if SUMMARY_STOP_LINE_RE.match(line):
            flush()
            break

        if ITEM_START_RE.match(line):
            flush()
            tokens = line.split()
            if len(tokens) >= 2:
                current = _parse_item_row_tokens(tokens)
                desc_parts = []

                need_unit = current.get("unit_price") is None
                need_total = current.get("total") is None
                if need_unit or need_total:
                    for j in (1, 2, 3, 4):
                        if i + j >= len(visual_lines):
                            break
                        nxt = visual_lines[i + j].strip()
                        if ITEM_START_RE.match(nxt) or SUMMARY_STOP_LINE_RE.match(nxt):
                            break
                        nxt_money = [t for t in nxt.split() if MONEY_RE.match(t)]
                        if nxt_money:
                            if need_total and current.get("total") is None:
                                current["total"] = nxt_money[-1]
                                need_total = False
                            if need_unit and len(nxt_money) >= 2 and current.get("unit_price") is None:
                                current["unit_price"] = nxt_money[-2]
                                need_unit = False
            i += 1
            continue

        if current:
            toks = line.split()
            wordy = any(
                not (QTY_RE.match(t) or MONEY_RE.match(t) or re.fullmatch(r"[A-Z/]+", t) or t.isdigit())
                for t in toks
            )
            if wordy and len(desc_parts) < 3:
                desc_parts.append(line)

        i += 1

    flush()

    cleaned: List[Dict[str, str]] = []
    for it in items:
        if not it.get("item_id"):
            continue
        cleaned.append(
            {
                "line_no": str(it.get("line_no") or ""),
                "item_id": str(it.get("item_id") or ""),
                "qty": str(it.get("qty") or ""),
                "unit_price": str(it.get("unit_price") or ""),
                "total": str(it.get("total") or ""),
                "description": str(it.get("description") or ""),
            }
        )
    return cleaned


def extract_line_items_hybrid(pdf_bytes: bytes) -> List[Dict[str, str]]:
    items = extract_line_items_by_tables(pdf_bytes)
    if not items:
        items = extract_line_items_by_words(pdf_bytes)

    seen = set()
    out = []
    for it in items:
        k = (it.get("line_no", ""), it.get("item_id", ""), it.get("total", ""))
        if k in seen:
            continue
        seen.add(k)
        out.append(it)
    return out


def build_rows_cadre(pdf_bytes: bytes, filename: str) -> List[Dict[str, Any]]:
    full_text = extract_full_text(pdf_bytes)
    header = extract_header_info_cadre(full_text)
    items = extract_line_items_hybrid(pdf_bytes)

    # Tax rule: 0 => ignore; >0 => include Tax row with blank desc
    tax_val = extract_tax_amount(full_text)
    if tax_val is not None and abs(tax_val) >= 0.005:
        tax_str = f"{tax_val:,.2f}"
        items.append({"line_no": "TAX", "item_id": "Tax", "qty": "1", "unit_price": tax_str, "total": tax_str, "description": ""})

    rows: List[Dict[str, Any]] = []
    for it in items:
        rows.append(
            {
                "ReferralManager": header.get("ReferralManager"),
                "ReferralEmail": None,
                "Brand": "Cadre Wire Group",
                "QuoteNumber": header.get("QuoteNumber"),
                "QuoteDate": normalize_date_str(header.get("QuoteDate")),
                "Company": header.get("Company"),
                "FirstName": header.get("FirstName"),
                "LastName": header.get("LastName"),
                "ContactEmail": None,
                "ContactPhone": None,
                "Address": header.get("Address"),
                "County": None,
                "City": header.get("City"),
                "State": header.get("State"),
                "ZipCode": header.get("ZipCode"),
                "Country": header.get("Country"),
                "item_id": it.get("item_id"),
                "item_desc": it.get("description"),
                "UnitPrice": _safe_float(it.get("unit_price")),
                "TotalSales": _safe_float(it.get("total")),
                "QuoteValidDate": normalize_date_str(header.get("QuoteValidDate")),
                "CustomerNumber": header.get("CustomerNumber"),
                "manufacturer_Name": None,
                "PDF": filename,
                "DemoQuote": None,
            }
        )
    return rows


# =========================================================
# VOELKR EXTRACTION (HEADER-FIELD MAPPING IN MEMORY)
# =========================================================
def build_rows_voelkr(pdf_bytes: bytes, filename: str) -> List[Dict[str, Any]]:
    """
    Produces exactly ONE row per PDF using VOELKR_FIELDS.
    Mapping is embedded in VOELKR_REGEX_MAP (in memory).
    """
    text = extract_full_text(pdf_bytes)

    row: Dict[str, Any] = {c: "" for c in VOELKR_FIELDS}
    row["PDF"] = filename if "PDF" in row else filename  # ensure we keep PDF name somewhere

    for field in VOELKR_FIELDS:
        if field == "PDF":
            continue
        pattern = VOELKR_REGEX_MAP.get(field, "")
        if not pattern:
            continue
        try:
            m = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
            if m:
                value = m.group(1) if m.lastindex else m.group(0)
                value = re.sub(r"\s+", " ", str(value)).strip()
                # normalize date-like fields
                if "date" in field.lower():
                    value = normalize_date_str(value) or value
                row[field] = value
        except re.error:
            # bad regex => keep blank
            pass

    return [row]


# =========================================================
# ROUTER
# =========================================================
def parse_pdf(system_type: str, pdf_bytes: bytes, filename: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    warnings: List[str] = []
    if system_type == "Cadre":
        return build_rows_cadre(pdf_bytes, filename), warnings
    if system_type == "Voelkr":
        # sanity check: you must paste your real headers
        if len(VOELKR_FIELDS) <= 3:
            warnings.append("Voelkr headers look like placeholders. Replace VOELKR_FIELDS with your real CSV headers.")
        return build_rows_voelkr(pdf_bytes, filename), warnings
    raise ValueError(f"Unknown system type: {system_type}")


def columns_for_system(system_type: str) -> List[str]:
    if system_type == "Cadre":
        return CADRE_COLUMNS
    if system_type == "Voelkr":
        return VOELKR_FIELDS
    return []


# =========================================================
# STREAMLIT UI
# =========================================================
st.set_page_config(page_title="PDF Extractor", layout="wide")
st.title("PDF Extractor")

st.markdown(
    """
- Select **System type** (Cadre / Voelkr)
- Upload PDF(s)
- Click **Extract**
- Download **one ZIP** (CSV + PDFs)

**Privacy:** Files are processed in-memory and not stored to disk by this app.
"""
)

# session state for clearing uploader & storing result zip
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "result_zip_bytes" not in st.session_state:
    st.session_state["result_zip_bytes"] = None
if "result_summary" not in st.session_state:
    st.session_state["result_summary"] = None
if "result_preview" not in st.session_state:
    st.session_state["result_preview"] = None

system_type = st.selectbox("System type", SYSTEM_TYPES, index=0)
debug_mode = st.checkbox("Debug mode", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['uploader_key']}",
)

extract_btn = st.button("Extract")

# show results if present
if st.session_state["result_zip_bytes"] is not None:
    st.success(st.session_state["result_summary"] or "Done.")
    if st.session_state["result_preview"] is not None:
        st.subheader("Preview (first 50 rows)")
        st.dataframe(st.session_state["result_preview"], use_container_width=True)

    st.download_button(
        "Download ZIP (CSV + PDFs)",
        data=st.session_state["result_zip_bytes"],
        file_name=f"{system_type.lower()}_extraction_output.zip",
        mime="application/zip",
    )

    if st.button("New extraction"):
        st.session_state["result_zip_bytes"] = None
        st.session_state["result_summary"] = None
        st.session_state["result_preview"] = None
        st.session_state["uploader_key"] += 1
        st.rerun()

if extract_btn:
    if not uploaded_files:
        st.error("Please upload at least one PDF.")
    else:
        file_data = [{"name": f.name, "bytes": f.read()} for f in uploaded_files]

        all_rows: List[Dict[str, Any]] = []
        warnings_all: List[str] = []

        progress = st.progress(0.0)
        status = st.empty()

        for idx, fd in enumerate(file_data, start=1):
            status.text(f"Processing {idx}/{len(file_data)}: {fd['name']}")
            rows, warns = parse_pdf(system_type, fd["bytes"], fd["name"])
            all_rows.extend(rows)
            warnings_all.extend([f"{fd['name']}: {w}" for w in warns])
            progress.progress(idx / len(file_data))

        cols = columns_for_system(system_type)
        df = pd.DataFrame(all_rows)
        if cols:
            # ensure all expected columns exist
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[cols]

        # build one ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("extracted/extracted.csv", df.to_csv(index=False).encode("utf-8"))
            for fd in file_data:
                zf.writestr(f"pdfs/{fd['name']}", fd["bytes"])

            if warnings_all:
                zf.writestr("extracted/warnings.txt", "\n".join(warnings_all).encode("utf-8"))

        zip_buf.seek(0)

        st.session_state["result_zip_bytes"] = zip_buf.getvalue()
        st.session_state["result_summary"] = f"Parsed {len(file_data)} PDF(s). Output rows: {len(df)}."
        st.session_state["result_preview"] = df.head(50)

        # clear uploader after extraction
        st.session_state["uploader_key"] += 1

        if debug_mode and warnings_all:
            st.warning("Warnings were generated. They are included inside the ZIP as extracted/warnings.txt")

        st.rerun()
