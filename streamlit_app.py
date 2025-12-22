import io
import re
import zipfile
from typing import List, Dict, Any
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st


# =========================================================
# SYSTEMS
# =========================================================
SYSTEM_TYPES = ["Cadre", "Voelkr"]  # Cadre stays blank for now


# =========================================================
# OUTPUT HEADERS (EXACT)
# =========================================================
VOELKR_COLUMNS = [
    "ReferralManager",
    "ReferralEmail",
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
    "manufacturer_Name",
    "item_id",
    "item_desc",
    "Quantity",
    "TotalSales",
    "PDF",
    "Brand",
    "QuoteExpiration",
    "CustomerNumber",
    "UnitSales",
    "Unit_Cost",
    "sales_cost",
    "cust_type",
    "QuoteComment",
    "Created_By",
    "quote_line_no",
    "DemoQuote",
]

VOELKR_DEFAULT_BRAND = "Voelker Controls"


# =========================================================
# NORMALIZERS
# =========================================================
def normalize_date(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})", s)
    if not m:
        return s.strip()
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def normalize_money(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"[\d,]+\.\d{2}", s)
    return m.group(0) if m else s


# =========================================================
# PDF EXTRACTION
# =========================================================
def extract_words_all_pages(pdf_bytes: bytes) -> List[dict]:
    words_all: List[dict] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=True,
            ) or []
            for w in words:
                w["page_num"] = page.page_number
            words_all.extend(words)
    return words_all


def group_words_into_lines(words: List[dict], y_tol=3.0, page_num=1) -> List[str]:
    buckets = defaultdict(list)
    for w in words:
        if w.get("page_num") != page_num:
            continue
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)

    lines: List[str] = []
    for y in sorted(buckets.keys()):
        row = sorted(buckets[y], key=lambda x: x["x0"])
        line = " ".join(w["text"] for w in row).strip()
        if line:
            lines.append(line)
    return lines


def extract_full_text(pdf_bytes: bytes) -> str:
    parts: List[str] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


# =========================================================
# VOELKR EXTRACTION (ROBUST)
# =========================================================
DATE_RE = re.compile(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b")
QUOTE_NO_RE = re.compile(r"\b(\d{2}/\d{6})\b")
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
USA_RE = re.compile(r"\b(USA|United States of America)\b", re.IGNORECASE)
MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


def extract_quote_date(lines: List[str], full_text: str, quote_number: str) -> str:
    for ln in lines:
        if "date" in ln.lower():
            m = DATE_RE.search(ln)
            if m:
                return normalize_date(m.group(1))

    if quote_number:
        for i, ln in enumerate(lines):
            if quote_number in ln:
                for j in range(i, min(i + 10, len(lines))):
                    m = DATE_RE.search(lines[j])
                    if m:
                        return normalize_date(m.group(1))

    m = DATE_RE.search(full_text)
    return normalize_date(m.group(1)) if m else ""


def extract_customer_number_from_words(words: List[dict]) -> str:
    if not words:
        return ""

    w1 = [w for w in words if w.get("page_num") == 1]

    def row_key(w, y_tol=3.0):
        return round(w["top"] / y_tol) * y_tol

    rows = defaultdict(list)
    for w in w1:
        rows[row_key(w)].append(w)

    sorted_row_keys = sorted(rows.keys())
    num_re = re.compile(r"^\d{3,10}$")

    # Find 'Cust' label and nearest numeric to right / below
    for idx, rk in enumerate(sorted_row_keys):
        row = sorted(rows[rk], key=lambda x: x["x0"])
        for w in row:
            t = (w.get("text") or "").strip()
            if not t:
                continue
            if t.lower().startswith("cust") or t.lower().startswith("customer"):
                label_x = w["x0"]

                # same row right side
                right_words = sorted([ww for ww in row if ww["x0"] > label_x], key=lambda x: x["x0"])
                for ww in right_words[:20]:
                    val = (ww.get("text") or "").strip()
                    val = re.sub(r"[^\d]", "", val)
                    if num_re.match(val):
                        return val

                # next rows (1-3)
                for down in (1, 2, 3):
                    if idx + down >= len(sorted_row_keys):
                        break
                    row2 = sorted(rows[sorted_row_keys[idx + down]], key=lambda x: x["x0"])
                    nearby = sorted([ww for ww in row2 if ww["x0"] >= label_x - 10], key=lambda x: x["x0"])
                    for ww in nearby[:20]:
                        val = (ww.get("text") or "").strip()
                        val = re.sub(r"[^\d]", "", val)
                        if num_re.match(val):
                            return val

    return ""


def extract_created_by(lines: List[str], full_text: str) -> str:
    """
    Created_By was missing: some PDFs show it as 'Created By', 'Prepared By',
    'Created:' or even just a name near a 'Created' label.
    We'll try:
      1) Line-based label matches (most reliable on word-lines)
      2) Full text label matches
      3) Fallback: look for 'Created' line and capture trailing name
    """
    patterns = [
        re.compile(r"\bCreated\s*By\b\s*[:#]?\s*(.+)$", re.IGNORECASE),
        re.compile(r"\bPrepared\s*By\b\s*[:#]?\s*(.+)$", re.IGNORECASE),
        re.compile(r"\bEntered\s*By\b\s*[:#]?\s*(.+)$", re.IGNORECASE),
        re.compile(r"\bCreated\b\s*[:#]?\s*(.+)$", re.IGNORECASE),
    ]

    for ln in lines:
        for pat in patterns:
            m = pat.search(ln)
            if m:
                val = m.group(1).strip()
                val = re.sub(r"\s{2,}", " ", val)
                # trim obvious trailing tokens
                val = re.split(r"\b(Date|Quote|Customer|Ship|Bill)\b", val, flags=re.IGNORECASE)[0].strip()
                if val:
                    return val

    for pat in patterns:
        m = pat.search(full_text)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"\s{2,}", " ", val)
            val = re.split(r"\b(Date|Quote|Customer|Ship|Bill)\b", val, flags=re.IGNORECASE)[0].strip()
            if val:
                return val

    return ""


def extract_ship_to_block(lines: List[str]) -> Dict[str, str]:
    """
    USER REQUIREMENT:
      Address must come from SHIP TO (not bill to / sold to).

    We find the 'Ship To' line and then read the next lines:
      Ship To
      COMPANY
      STREET
      CITY STATE ZIP
      USA
    """
    out: Dict[str, str] = {}

    # locate Ship To marker
    ship_idx = -1
    for i, ln in enumerate(lines):
        if re.search(r"\bShip\s*To\b", ln, re.IGNORECASE):
            ship_idx = i
            break

    if ship_idx == -1:
        return out

    # gather next 6-8 lines after Ship To
    block = []
    for j in range(ship_idx + 1, min(ship_idx + 10, len(lines))):
        t = lines[j].strip()
        if not t:
            continue
        # stop if we hit another header section
        if re.search(r"\b(Bill\s*To|Sold\s*To|Quote|Items?|Subtotal|Total)\b", t, re.IGNORECASE):
            break
        block.append(t)

    if not block:
        return out

    # remove duplicate "Ship To:" content if embedded
    if block and re.search(r"\bShip\s*To\b", block[0], re.IGNORECASE):
        block = block[1:]

    # Heuristic parse
    # company = first non-empty line that isn't city/state/zip
    # street = next line that starts with digits OR looks like street
    # city/state/zip = line containing ZIP
    company = block[0] if block else ""
    if company:
        out["Company"] = company

    # find city/state/zip line
    city_state_zip_line = ""
    city_state_zip_idx = -1
    for k, ln in enumerate(block):
        if ZIP_RE.search(ln):
            city_state_zip_line = ln
            city_state_zip_idx = k
            break

    # street line: typically line just before city/state/zip
    if city_state_zip_idx > 0:
        out["Address"] = block[city_state_zip_idx - 1].strip()

    # parse city/state/zip
    if city_state_zip_line:
        zip_code = ZIP_RE.search(city_state_zip_line).group(1)
        out["ZipCode"] = zip_code

        tokens = city_state_zip_line.split()
        state = ""
        if zip_code in tokens:
            zidx = tokens.index(zip_code)
            if zidx >= 1 and re.fullmatch(r"[A-Z]{2}", tokens[zidx - 1]):
                state = tokens[zidx - 1]
        if state:
            out["State"] = state
            sidx = tokens.index(state)
            if sidx >= 1:
                out["City"] = " ".join(tokens[:sidx]).strip()
        else:
            # fallback: last 2-letter token before zip
            for t in reversed(tokens):
                if re.fullmatch(r"[A-Z]{2}", t):
                    out["State"] = t
                    break

    # country: search within block
    out["Country"] = "USA"
    for ln in block:
        if USA_RE.search(ln):
            out["Country"] = "USA"
            break

    return out


def extract_total_sales(lines: List[str], full_text: str) -> str:
    total = ""
    for ln in lines:
        if re.search(r"\b(total|grand total|total sales)\b", ln, re.IGNORECASE):
            mm = MONEY_RE.search(ln)
            if mm:
                total = mm.group(0)
                break
    if not total:
        m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})",
                      full_text, flags=re.IGNORECASE)
        if m:
            total = m.group(1)
        else:
            monies = MONEY_RE.findall(full_text)
            if monies:
                total = monies[-1]
    return normalize_money(total) if total else ""


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    words = extract_words_all_pages(pdf_bytes)
    lines = group_words_into_lines(words, y_tol=3.0, page_num=1)
    full_text = extract_full_text(pdf_bytes)

    out: Dict[str, str] = {}

    # QuoteNumber
    m = QUOTE_NO_RE.search(full_text)
    if m:
        out["QuoteNumber"] = m.group(1).strip()

    # QuoteDate
    out["QuoteDate"] = extract_quote_date(lines, full_text, out.get("QuoteNumber", ""))

    # CustomerNumber
    out["CustomerNumber"] = extract_customer_number_from_words(words)

    # ReferralManager (sample: DAYTON)
    if any(re.search(r"\bDAYTON\b", ln) for ln in lines) or re.search(r"\bDAYTON\b", full_text):
        out["ReferralManager"] = "DAYTON"

    # TotalSales
    out["TotalSales"] = extract_total_sales(lines, full_text)

    # Created_By (yellow highlight)
    out["Created_By"] = extract_created_by(lines, full_text)

    # SHIP TO address block (red highlight)
    out.update(extract_ship_to_block(lines))

    # Brand default
    out["Brand"] = VOELKR_DEFAULT_BRAND

    return out


def build_voelkr_row(pdf_bytes: bytes, filename: str) -> Dict[str, Any]:
    extracted = extract_voelkr_fields(pdf_bytes)
    row = {c: "" for c in VOELKR_COLUMNS}
    row.update(extracted)
    row["PDF"] = filename
    row["Brand"] = VOELKR_DEFAULT_BRAND
    return row


# =========================================================
# STREAMLIT UI
# =========================================================
st.set_page_config(page_title="PDF Extractor", layout="wide")
st.title("PDF Extractor – Voelkr")

st.markdown(
    """
- **Voelkr** extracts header fields (Ship To address + Created_By).
- **Cadre** is blank for now.
- Download is **one ZIP** containing extracted CSV + uploaded PDFs.
"""
)

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "result_zip" not in st.session_state:
    st.session_state["result_zip"] = None
if "result_df" not in st.session_state:
    st.session_state["result_df"] = None
if "result_summary" not in st.session_state:
    st.session_state["result_summary"] = None

system_type = st.selectbox("System type", SYSTEM_TYPES, index=1)

uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['uploader_key']}",
)

extract_btn = st.button("Extract")

if st.session_state["result_zip"] is not None:
    st.success(st.session_state["result_summary"] or "Done.")
    st.dataframe(st.session_state["result_df"].head(50), use_container_width=True)

    st.download_button(
        "Download ZIP (CSV + PDFs)",
        data=st.session_state["result_zip"],
        file_name=f"{system_type.lower()}_output.zip",
        mime="application/zip",
    )

    if st.button("New extraction"):
        st.session_state["result_zip"] = None
        st.session_state["result_df"] = None
        st.session_state["result_summary"] = None
        st.session_state["uploader_key"] += 1
        st.rerun()

if extract_btn:
    if system_type == "Cadre":
        st.info("Cadre mapping is not configured yet. Select **Voelkr**.")
    else:
        if not uploaded_files:
            st.error("Please upload at least one PDF.")
        else:
            file_data = [{"name": f.name, "bytes": f.read()} for f in uploaded_files]

            rows: List[Dict[str, Any]] = []
            progress = st.progress(0.0)
            status = st.empty()

            for idx, fd in enumerate(file_data, start=1):
                status.text(f"Processing {idx}/{len(file_data)}: {fd['name']}")
                rows.append(build_voelkr_row(fd["bytes"], fd["name"]))
                progress.progress(idx / len(file_data))

            df = pd.DataFrame(rows, columns=VOELKR_COLUMNS)

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("extracted/voelkr_extracted.csv", df.to_csv(index=False).encode("utf-8"))
                for fd in file_data:
                    zf.writestr(f"pdfs/{fd['name']}", fd["bytes"])
            zip_buf.seek(0)

            st.session_state["result_zip"] = zip_buf.getvalue()
            st.session_state["result_df"] = df
            st.session_state["result_summary"] = f"Parsed {len(file_data)} PDF(s). Output rows: {len(df)}."

            st.session_state["uploader_key"] += 1
            st.rerun()
