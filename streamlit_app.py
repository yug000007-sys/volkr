import io
import re
import zipfile
from typing import List, Dict, Any, Optional

import pdfplumber
import pandas as pd
import streamlit as st


# =========================================================
# SYSTEMS
# =========================================================
SYSTEM_TYPES = ["Cadre", "Voelkr"]  # Cadre stays blank for now


# =========================================================
# OUTPUT HEADERS (EXACTLY as you provided)
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
# PDF TEXT EXTRACTION
# =========================================================
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


# =========================================================
# VOELKR: BETTER QUOTEDATE + ADDRESS BLOCK EXTRACTION
# =========================================================
def extract_quote_date(full_text: str, quote_number: str) -> str:
    """
    QuoteDate was missing for you.
    We do:
      1) If 'Date' label exists, prefer Date: mm/dd/yyyy
      2) If QuoteNumber exists, look within 500 chars AFTER it for first date
      3) Otherwise, take the first date in the document
    """
    # 1) labeled date (highest precision)
    m = re.search(r"\bDate\b\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{4})", full_text, flags=re.IGNORECASE)
    if m:
        return normalize_date(m.group(1))

    # 2) near quote number
    if quote_number:
        qn = re.escape(quote_number)
        m = re.search(qn + r"(.{0,500}?)(\d{1,2}/\d{1,2}/\d{4})", full_text, flags=re.DOTALL)
        if m:
            return normalize_date(m.group(2))

    # 3) first date anywhere
    m = re.search(r"\b(\d{1,2}/\d{1,2}/\d{4})\b", full_text)
    if m:
        return normalize_date(m.group(1))

    return ""


def extract_company_address_city_state_zip_country(full_text: str) -> Dict[str, str]:
    """
    Address block was missing. We'll extract:
      - Company (line above street when possible)
      - Street address
      - City, State, Zip
      - Country (USA)

    Strategy:
      A) Strong: find a US address block with (street + city/state/zip)
         Works with normal line breaks.
      B) Fallback: same, but allow everything on ONE line (some PDFs flatten lines).
    """
    out: Dict[str, str] = {}

    # A) Multi-line address block
    m = re.search(
        r"(?P<company>[A-Z0-9][A-Z0-9 &',.\-]{3,})\s*\n+"
        r"(?P<street>\d{3,6}\s+[A-Z0-9][A-Z0-9 .#&'\-]+)\s*\n+"
        r"(?P<city>[A-Z][A-Z .'\-]+?)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\s*\n+"
        r"(?P<country>USA|United States of America)\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        out["Company"] = m.group("company").strip()
        out["Address"] = m.group("street").strip()
        out["City"] = m.group("city").strip()
        out["State"] = m.group("state").strip().upper()
        out["ZipCode"] = m.group("zip").strip()
        out["Country"] = "USA"
        return out

    # B) Fallback when PDF text is flattened onto one line
    # Example: "SINBON USA LLC 4265 GIBSON DRIVE TIPP CITY OH 45371 USA"
    m = re.search(
        r"(?P<company>[A-Z0-9][A-Z0-9 &',.\-]{3,})\s+"
        r"(?P<street>\d{3,6}\s+[A-Z0-9][A-Z0-9 .#&'\-]+)\s+"
        r"(?P<city>[A-Z][A-Z .'\-]+?)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\s+"
        r"(?P<country>USA|United States of America)\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        out["Company"] = m.group("company").strip()
        out["Address"] = m.group("street").strip()
        out["City"] = m.group("city").strip()
        out["State"] = m.group("state").strip().upper()
        out["ZipCode"] = m.group("zip").strip()
        out["Country"] = "USA"
        return out

    # C) If company is not in the same block, still try to get street + city/state/zip
    m = re.search(
        r"(?P<street>\d{3,6}\s+[A-Z0-9][A-Z0-9 .#&'\-]+)\s*\n+"
        r"(?P<city>[A-Z][A-Z .'\-]+?)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        out["Address"] = m.group("street").strip()
        out["City"] = m.group("city").strip()
        out["State"] = m.group("state").strip().upper()
        out["ZipCode"] = m.group("zip").strip()
        # Country: if USA appears anywhere, set it
        if re.search(r"\b(USA|United States of America)\b", full_text, flags=re.IGNORECASE):
            out["Country"] = "USA"

        # best-effort company: nearest non-empty line above street
        street_line = m.group("street").strip()
        pos = full_text.lower().find(street_line.lower())
        if pos != -1:
            before = full_text[max(0, pos - 250):pos]
            lines = [ln.strip() for ln in before.splitlines() if ln.strip()]
            for cand in reversed(lines[-8:]):
                if len(cand) >= 4 and not re.search(r"(phone|fax|date|quote|customer|ship|bill|sold\s*to)", cand, re.IGNORECASE):
                    if not re.match(r"^\d{3,6}\s", cand):
                        out["Company"] = cand
                        break

    return out


def extract_voelkr_fields(full_text: str) -> Dict[str, str]:
    out: Dict[str, str] = {}

    # QuoteNumber: example "01/125785"
    m = re.search(r"\b(\d{2}/\d{6})\b", full_text)
    if m:
        out["QuoteNumber"] = m.group(1).strip()

    # QuoteDate: improved
    out["QuoteDate"] = extract_quote_date(full_text, out.get("QuoteNumber", ""))

    # CustomerNumber: example 10980
    m = re.search(r"Customer\s*(?:No\.?|Number)?\s*[:#]?\s*(\d{3,10})", full_text, flags=re.IGNORECASE)
    if m:
        out["CustomerNumber"] = m.group(1).strip()

    # ReferralManager: DAYTON
    m = re.search(r"\bDAYTON\b", full_text)
    if m:
        out["ReferralManager"] = "DAYTON"

    # TotalSales
    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["TotalSales"] = normalize_money(m.group(1))
    else:
        monies = re.findall(r"\b[\d,]+\.\d{2}\b", full_text)
        if monies:
            out["TotalSales"] = normalize_money(monies[-1])

    # Created_By
    m = re.search(r"(?:Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Za-z][A-Za-z .'\-]+)",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["Created_By"] = m.group(1).strip()
    else:
        m = re.search(r"\bLoghan\s+Keefer\b", full_text, flags=re.IGNORECASE)
        if m:
            out["Created_By"] = "Loghan Keefer"

    # Address block: improved (Company, Address, City, State, Zip, Country)
    out.update(extract_company_address_city_state_zip_country(full_text))

    # Brand default always
    out["Brand"] = VOELKR_DEFAULT_BRAND

    return out


def build_voelkr_row(pdf_bytes: bytes, filename: str) -> Dict[str, Any]:
    full_text = extract_full_text(pdf_bytes)
    extracted = extract_voelkr_fields(full_text)

    row = {c: "" for c in VOELKR_COLUMNS}
    row.update(extracted)
    row["PDF"] = filename
    row["Brand"] = VOELKR_DEFAULT_BRAND

    # normalize
    if row.get("QuoteDate"):
        row["QuoteDate"] = normalize_date(row["QuoteDate"])
    if row.get("TotalSales"):
        row["TotalSales"] = normalize_money(row["TotalSales"])
    if row.get("Country"):
        if row["Country"].strip().lower().startswith("united") or row["Country"].strip().upper() == "USA":
            row["Country"] = "USA"

    return row


# =========================================================
# STREAMLIT UI
# =========================================================
st.set_page_config(page_title="PDF Extractor", layout="wide")
st.title("PDF Extractor")

st.markdown(
    """
Select **Voelkr** to extract header fields.

- **Cadre** is intentionally blank for now (mapping will be added later).
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

            # build ZIP: CSV + PDFs
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("extracted/voelkr_extracted.csv", df.to_csv(index=False).encode("utf-8"))
                for fd in file_data:
                    zf.writestr(f"pdfs/{fd['name']}", fd["bytes"])
            zip_buf.seek(0)

            st.session_state["result_zip"] = zip_buf.getvalue()
            st.session_state["result_df"] = df
            st.session_state["result_summary"] = f"Parsed {len(file_data)} PDF(s). Output rows: {len(df)}."

            # clear uploader
            st.session_state["uploader_key"] += 1
            st.rerun()
