import io
import re
import zipfile
from typing import List, Dict, Any, Optional
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st


SYSTEM_TYPES = ["Cadre", "Voelkr"]

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


# -------------------------
# BASIC NORMALIZERS
# -------------------------
def normalize_date(s: str) -> str:
    s = (s or "").strip()
    m = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})", s)
    if not m:
        return s
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def normalize_money(s: str) -> str:
    s = (s or "").strip()
    m = re.search(r"[\d,]+\.\d{2}", s)
    return m.group(0) if m else s


# -------------------------
# WORD-BASED LINE BUILDER
# -------------------------
def extract_words_all_pages(pdf_bytes: bytes) -> List[dict]:
    words_all: List[dict] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # "extract_words" is usually more reliable than extract_text for structured headers
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


def group_words_into_lines(words: List[dict], y_tol: float = 3.0, only_first_page: bool = True) -> List[str]:
    """
    Groups words into visual lines by y coordinate.
    """
    buckets = defaultdict(list)
    for w in words:
        if only_first_page and w.get("page_num") != 1:
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
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                t = page.extract_text(layout=True) or ""
            except TypeError:
                t = page.extract_text() or ""
            parts.append(t)
    return "\n".join(parts)


# -------------------------
# VOELKR EXTRACTION (ROBUST)
# -------------------------
DATE_RE = re.compile(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b")
QUOTE_NO_RE = re.compile(r"\b(\d{2}/\d{6})\b")
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
STATE_RE = re.compile(r"\b([A-Z]{2})\b")
USA_RE = re.compile(r"\b(USA|United States of America)\b", re.IGNORECASE)

MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


def find_first_match(lines: List[str], pattern: re.Pattern, prefer_label: Optional[str] = None) -> str:
    """
    Find match in lines.
    If prefer_label is provided, first look in lines that contain that label.
    """
    if prefer_label:
        for ln in lines:
            if prefer_label.lower() in ln.lower():
                m = pattern.search(ln)
                if m:
                    return m.group(1)
    for ln in lines:
        m = pattern.search(ln)
        if m:
            return m.group(1)
    return ""


def extract_address_block_from_lines(lines: List[str]) -> Dict[str, str]:
    """
    Extract Company, Address, City, State, Zip, Country from "visual lines".
    Works when PDF text is broken in extract_text().
    """
    out: Dict[str, str] = {}

    # We locate a line containing a ZIP. Usually address appears as:
    #   COMPANY
    #   4265 GIBSON DRIVE
    #   TIPP CITY OH 45371
    #   USA
    #
    # We'll scan for a "city state zip" line, then take street line above it,
    # company line above street, and optional country line below.
    for i, ln in enumerate(lines):
        if not ZIP_RE.search(ln):
            continue

        # Try parse city/state/zip from this line
        # e.g., "TIPP CITY OH 45371"
        m_zip = ZIP_RE.search(ln)
        zip_code = m_zip.group(1) if m_zip else ""
        # state = last 2-letter token before zip (best effort)
        tokens = ln.split()
        state = ""
        if zip_code and zip_code in tokens:
            zidx = tokens.index(zip_code)
            if zidx >= 1:
                cand = tokens[zidx - 1]
                if re.fullmatch(r"[A-Z]{2}", cand):
                    state = cand

        # city = tokens before state (join)
        city = ""
        if state and state in tokens:
            sidx = tokens.index(state)
            if sidx >= 1:
                city = " ".join(tokens[:sidx]).strip()

        if zip_code:
            out["ZipCode"] = zip_code
        if state:
            out["State"] = state
        if city:
            out["City"] = city

        # street = previous non-empty line
        street = ""
        j = i - 1
        while j >= 0:
            if lines[j].strip():
                street = lines[j].strip()
                break
            j -= 1
        if street:
            out["Address"] = street

        # company = line above street
        company = ""
        if j - 1 >= 0:
            company = lines[j - 1].strip()
        # company sanity: avoid labels
        if company and not re.search(r"(date|quote|customer|phone|fax|ship|bill|sold\s*to)", company, re.IGNORECASE):
            out["Company"] = company

        # country = next line if it contains USA
        country = ""
        if i + 1 < len(lines) and USA_RE.search(lines[i + 1]):
            country = "USA"
        elif USA_RE.search(ln):
            country = "USA"
        else:
            # sometimes USA appears elsewhere in header
            for k in range(i, min(i + 4, len(lines))):
                if USA_RE.search(lines[k]):
                    country = "USA"
                    break

        if country:
            out["Country"] = "USA"

        # If we got zip + address, we consider it found and stop.
        if out.get("ZipCode") and out.get("Address"):
            break

    return out


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    """
    Primary extraction uses word-built lines (page 1),
    fallback to full text if needed.
    """
    out: Dict[str, str] = {}

    words = extract_words_all_pages(pdf_bytes)
    lines = group_words_into_lines(words, y_tol=3.0, only_first_page=True)

    # Fallback raw text
    full_text = extract_full_text(pdf_bytes)

    # QuoteNumber
    qn = find_first_match(lines, QUOTE_NO_RE) or (QUOTE_NO_RE.search(full_text).group(1) if QUOTE_NO_RE.search(full_text) else "")
    out["QuoteNumber"] = qn

    # QuoteDate (prefer lines containing "Date")
    qd = find_first_match(lines, DATE_RE, prefer_label="Date")
    if not qd and qn:
        # look in nearby lines after quote number
        for i, ln in enumerate(lines):
            if qn in ln:
                for j in range(i, min(i + 8, len(lines))):
                    m = DATE_RE.search(lines[j])
                    if m:
                        qd = m.group(1)
                        break
            if qd:
                break
    if not qd:
        # fallback: first date in full text
        m = DATE_RE.search(full_text)
        qd = m.group(1) if m else ""
    out["QuoteDate"] = normalize_date(qd) if qd else ""

    # ReferralManager (your sample: DAYTON)
    if any(re.search(r"\bDAYTON\b", ln) for ln in lines) or re.search(r"\bDAYTON\b", full_text):
        out["ReferralManager"] = "DAYTON"

    # CustomerNumber
    m = re.search(r"Customer\s*(?:No\.?|Number)?\s*[:#]?\s*(\d{3,10})", full_text, flags=re.IGNORECASE)
    if m:
        out["CustomerNumber"] = m.group(1).strip()

    # TotalSales (prefer label lines)
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
    out["TotalSales"] = normalize_money(total) if total else ""

    # Created_By
    m = re.search(r"(?:Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Za-z][A-Za-z .'\-]+)",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["Created_By"] = m.group(1).strip()
    else:
        m = re.search(r"\bLoghan\s+Keefer\b", full_text, flags=re.IGNORECASE)
        if m:
            out["Created_By"] = "Loghan Keefer"

    # Address block (Company/Address/City/State/Zip/Country) from lines
    out.update(extract_address_block_from_lines(lines))

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
st.title("PDF Extractor")

st.markdown(
    """
Select **Voelkr** to extract header fields.

- **Cadre** is intentionally blank for now.
- Upload PDFs → Extract → Download one ZIP (CSV + PDFs)
- Files are processed **in-memory** (not saved to disk by the app).
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

system_type = st.selectbox("System type", SYSTEM_TYPES, index=1)  # default Voelkr
debug_mode = st.checkbox("Debug mode", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['uploader_key']}",
)

extract_btn = st.button("Extract")

# show results if present
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
        st.info("Cadre mapping is not configured yet. Select **Voelkr** to extract Voelkr PDFs.")
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
                row = build_voelkr_row(fd["bytes"], fd["name"])
                rows.append(row)

                if debug_mode:
                    st.write(
                        fd["name"],
                        {
                            "QuoteNumber": row.get("QuoteNumber", ""),
                            "QuoteDate": row.get("QuoteDate", ""),
                            "Company": row.get("Company", ""),
                            "Address": row.get("Address", ""),
                            "City": row.get("City", ""),
                            "State": row.get("State", ""),
                            "ZipCode": row.get("ZipCode", ""),
                            "Country": row.get("Country", ""),
                        },
                    )

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
