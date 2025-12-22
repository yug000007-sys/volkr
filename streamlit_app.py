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
SYSTEM_TYPES = ["Cadre", "Voelkr"]


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
# WORD-BASED EXTRACTION
# =========================================================
def extract_words_all_pages(pdf_bytes: bytes) -> List[dict]:
    words_all = []
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


def group_words_into_lines(words: List[dict], y_tol=3.0) -> List[str]:
    buckets = defaultdict(list)
    for w in words:
        if w.get("page_num") != 1:
            continue
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)

    lines = []
    for y in sorted(buckets.keys()):
        row = sorted(buckets[y], key=lambda x: x["x0"])
        line = " ".join(w["text"] for w in row).strip()
        if line:
            lines.append(line)
    return lines


def extract_full_text(pdf_bytes: bytes) -> str:
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            out.append(page.extract_text() or "")
    return "\n".join(out)


# =========================================================
# VOELKR FIELD EXTRACTION
# =========================================================
DATE_RE = re.compile(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}\b")
QUOTE_RE = re.compile(r"\b\d{2}/\d{6}\b")
ZIP_RE = re.compile(r"\b\d{5}\b")
MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


def extract_customer_number(lines: List[str], full_text: str) -> str:
    """
    Handles:
      Cust # 10980
      Cust#10980
      Customer # 10980
      Customer No: 10980
    """
    # 1️⃣ Prefer visual header lines
    for ln in lines:
        if re.search(r"\b(cust|customer)\b", ln, re.IGNORECASE):
            m = re.search(r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#]?\s*(\d{3,10})", ln, re.IGNORECASE)
            if m:
                return m.group(1)

    # 2️⃣ Fallback: full text
    m = re.search(
        r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#]?\s*(\d{3,10})",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return m.group(1)

    return ""


def extract_address_block(lines: List[str]) -> Dict[str, str]:
    out = {}

    for i, ln in enumerate(lines):
        if not ZIP_RE.search(ln):
            continue

        tokens = ln.split()
        zip_code = ZIP_RE.search(ln).group(0)
        state = tokens[tokens.index(zip_code) - 1] if zip_code in tokens else ""
        city = " ".join(tokens[: tokens.index(state)]) if state in tokens else ""

        out["ZipCode"] = zip_code
        out["State"] = state
        out["City"] = city

        if i - 1 >= 0:
            out["Address"] = lines[i - 1]
        if i - 2 >= 0:
            out["Company"] = lines[i - 2]

        out["Country"] = "USA"
        break

    return out


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    words = extract_words_all_pages(pdf_bytes)
    lines = group_words_into_lines(words)
    full_text = extract_full_text(pdf_bytes)

    out = {}

    # QuoteNumber
    m = QUOTE_RE.search(full_text)
    if m:
        out["QuoteNumber"] = m.group(0)

    # QuoteDate
    for ln in lines:
        if "date" in ln.lower():
            m = DATE_RE.search(ln)
            if m:
                out["QuoteDate"] = normalize_date(m.group(0))
                break
    if not out.get("QuoteDate"):
        m = DATE_RE.search(full_text)
        if m:
            out["QuoteDate"] = normalize_date(m.group(0))

    # CustomerNumber (FIXED)
    out["CustomerNumber"] = extract_customer_number(lines, full_text)

    # ReferralManager
    if any("DAYTON" in ln for ln in lines):
        out["ReferralManager"] = "DAYTON"

    # TotalSales
    for ln in lines:
        if "total" in ln.lower():
            m = MONEY_RE.search(ln)
            if m:
                out["TotalSales"] = normalize_money(m.group(0))
                break

    # Created_By
    m = re.search(r"(created|prepared)\s*by\s*[:#]?\s*([A-Za-z ]+)", full_text, re.IGNORECASE)
    if m:
        out["Created_By"] = m.group(2).strip()

    # Address block
    out.update(extract_address_block(lines))

    # Brand
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

system_type = st.selectbox("System type", SYSTEM_TYPES, index=1)
uploaded_files = st.file_uploader("Upload Voelkr PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Extract"):
    if system_type == "Cadre":
        st.info("Cadre mapping will be added later.")
    elif not uploaded_files:
        st.error("Upload at least one PDF.")
    else:
        rows = []
        for f in uploaded_files:
            rows.append(build_voelkr_row(f.read(), f.name))

        df = pd.DataFrame(rows, columns=VOELKR_COLUMNS)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("voelkr_extracted.csv", df.to_csv(index=False))
            for f in uploaded_files:
                zf.writestr(f.name, f.getvalue())

        zip_buf.seek(0)

        st.success(f"Extracted {len(df)} PDF(s)")
        st.dataframe(df, use_container_width=True)
        st.download_button("Download ZIP", zip_buf, "voelkr_output.zip")
