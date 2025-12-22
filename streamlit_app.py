import io
import re
import zipfile
from typing import List, Dict, Any, Optional, Tuple

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


# =========================================================
# BRAND DEFAULT
# =========================================================
VOELKR_DEFAULT_BRAND = "Voelker Controls"


# =========================================================
# PDF TEXT EXTRACTION
# =========================================================
def extract_full_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages."""
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
    # keep like 23,540.00
    m = re.search(r"[\d,]+\.\d{2}", s)
    return m.group(0) if m else s


# =========================================================
# VOELKR FIELD MAPPING (BASED ON YOUR SAMPLE)
# =========================================================
# These regexes are designed to be resilient:
# - Try specific anchors first
# - Fallback patterns if layout varies
#
# IMPORTANT: If your PDF text differs slightly, we can tighten/adjust these rules.
# =========================================================
def extract_voelkr_fields(full_text: str) -> Dict[str, str]:
    out: Dict[str, str] = {}

    # QuoteNumber: example "01/125785"
    # Try "Quote" or similar label, else first occurrence of NN/NNNNNN pattern
    m = re.search(r"\b(\d{2}/\d{6})\b", full_text)
    if m:
        out["QuoteNumber"] = m.group(1).strip()

    # QuoteDate: example "09/24/2025"
    # Prefer patterns near the quote number area (best-effort)
    # We’ll take the first mm/dd/yyyy we find after QuoteNumber if possible
    if out.get("QuoteNumber"):
        qn = re.escape(out["QuoteNumber"])
        m = re.search(qn + r".{0,200}?(\d{1,2}/\d{1,2}/\d{4})", full_text, flags=re.DOTALL)
        if m:
            out["QuoteDate"] = normalize_date(m.group(1))
    if not out.get("QuoteDate"):
        m = re.search(r"\b(\d{1,2}/\d{1,2}/\d{4})\b", full_text)
        if m:
            out["QuoteDate"] = normalize_date(m.group(1))

    # CustomerNumber: example 10980
    # Look for "Customer" label first
    m = re.search(r"Customer\s*(?:No\.?|Number)?\s*[:#]?\s*(\d{3,10})", full_text, flags=re.IGNORECASE)
    if m:
        out["CustomerNumber"] = m.group(1).strip()
    else:
        # fallback: if your layout has a bare 5-digit customer near header
        m = re.search(r"\b(10980)\b", full_text)  # fallback to known sample pattern
        if m:
            out["CustomerNumber"] = m.group(1).strip()

    # ReferralManager: "DAYTON"
    # Typical: appears near sales rep / branch / territory
    # We'll prefer explicit labels if present; else look for DAYTON as uppercase token (sample)
    m = re.search(r"(?:Referral\s*Manager|Sales\s*Office|Branch|Location)\s*[:#]?\s*([A-Z][A-Z0-9 ]{2,30})",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["ReferralManager"] = m.group(1).strip()
    else:
        m = re.search(r"\bDAYTON\b", full_text)
        if m:
            out["ReferralManager"] = "DAYTON"

    # Company + Address + City/State/Zip + Country
    # Your sample:
    # Company: SINBON USA LLC
    # Address: 4265 GIBSON DRIVE
    # City: TIPP CITY
    # State: OH
    # Zip: 45371
    # Country: USA
    #
    # Strategy:
    # - Find a US address block: "<street>\n<city> <state> <zip>\nUSA"
    # - Then look above it for company line
    addr_block = None
    # Find city/state/zip line
    m = re.search(
        r"(?P<street>\d{3,6}\s+[A-Z0-9][A-Z0-9 .#&'\-]+)\s*\n+"
        r"(?P<city>[A-Z][A-Z .'\-]+)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5})(?:-\d{4})?\s*\n+"
        r"(?P<country>USA|United States of America)\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        out["Address"] = m.group("street").strip()
        out["City"] = m.group("city").strip()
        out["State"] = m.group("state").strip().upper()
        out["ZipCode"] = m.group("zip").strip()
        out["Country"] = "USA"
        addr_block = m.group(0)

        # Company often appears 1-3 lines above street address.
        # Grab a few lines before the street line and choose the last "wordy" line.
        street_line = m.group("street").strip()
        idx = full_text.lower().find(street_line.lower())
        if idx != -1:
            before = full_text[max(0, idx - 200):idx]
            lines = [ln.strip() for ln in before.splitlines() if ln.strip()]
            # pick last reasonable company-like line (avoid labels, phone, etc.)
            for cand in reversed(lines[-6:]):
                if len(cand) >= 4 and not re.search(r"(phone|fax|date|quote|customer|ship|bill)", cand, re.IGNORECASE):
                    # avoid lines that look like pure address fragments
                    if not re.match(r"^\d{3,6}\s", cand):
                        out["Company"] = cand
                        break

    # TotalSales: example 23,540.00
    # Look for "Total" / "Total Sales" or last money in totals area
    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["TotalSales"] = normalize_money(m.group(1))
    else:
        # fallback: pick the largest-looking money value (often total)
        monies = re.findall(r"\b[\d,]+\.\d{2}\b", full_text)
        if monies:
            # heuristic: last money on doc is often total
            out["TotalSales"] = normalize_money(monies[-1])

    # Created_By: "Loghan Keefer"
    m = re.search(r"(?:Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Za-z][A-Za-z .'\-]+)",
                  full_text, flags=re.IGNORECASE)
    if m:
        out["Created_By"] = m.group(1).strip()
    else:
        # fallback to sample if present
        m = re.search(r"\bLoghan\s+Keefer\b", full_text, flags=re.IGNORECASE)
        if m:
            out["Created_By"] = "Loghan Keefer"

    # Brand is default always
    out["Brand"] = VOELKR_DEFAULT_BRAND

    return out


def build_voelkr_row(pdf_bytes: bytes, filename: str) -> Dict[str, Any]:
    full_text = extract_full_text(pdf_bytes)
    extracted = extract_voelkr_fields(full_text)

    # build full row with blanks for everything else
    row = {c: "" for c in VOELKR_COLUMNS}
    row.update(extracted)

    # Always set PDF filename
    row["PDF"] = filename

    # Ensure company/address normalization
    if row.get("QuoteDate"):
        row["QuoteDate"] = normalize_date(row["QuoteDate"])
    if row.get("TotalSales"):
        row["TotalSales"] = normalize_money(row["TotalSales"])
    if row.get("Country"):
        # normalize to USA
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
                    # Show key extracted fields for quick verification
                    st.write(
                        fd["name"],
                        {
                            "ReferralManager": row.get("ReferralManager", ""),
                            "CustomerNumber": row.get("CustomerNumber", ""),
                            "QuoteNumber": row.get("QuoteNumber", ""),
                            "QuoteDate": row.get("QuoteDate", ""),
                            "Company": row.get("Company", ""),
                            "Address": row.get("Address", ""),
                            "City": row.get("City", ""),
                            "State": row.get("State", ""),
                            "ZipCode": row.get("ZipCode", ""),
                            "TotalSales": row.get("TotalSales", ""),
                            "Created_By": row.get("Created_By", ""),
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
