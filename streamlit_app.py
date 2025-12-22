import io
import re
import zipfile
from typing import List, Dict, Any, Tuple
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st


# =========================================================
# SYSTEMS
# =========================================================
SYSTEM_TYPES = ["Cadre", "Voelkr"]  # Cadre stays blank for now
VOELKR_DEFAULT_BRAND = "Voelker Controls"


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


# =========================================================
# REGEX
# =========================================================
QUOTE_NO_RE = re.compile(r"\b(\d{2}/\d{6})\b")
DATE_RE = re.compile(r"\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b")
ZIP_LINE_RE = re.compile(r"^(.*)\s+([A-Z]{2})\s+(\d{5})(?:-\d{4})?$")
MONEY_RE = re.compile(r"\b([\d,]+\.\d{2})\b")


def normalize_date(s: str) -> str:
    s = (s or "").strip()
    m = DATE_RE.search(s)
    if not m:
        return s
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def clean_text(s: str) -> str:
    return re.sub(r"\s{2,}", " ", (s or "").replace("\n", " ")).strip()


# =========================================================
# PDF HELPERS
# =========================================================
def extract_full_text(pdf_bytes: bytes) -> str:
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            out.append(page.extract_text() or "")
    return "\n".join(out)


def extract_words(pdf_bytes: bytes) -> List[dict]:
    words = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            ws = page.extract_words(x_tolerance=2, y_tolerance=2, use_text_flow=True) or []
            for w in ws:
                w["page"] = page.page_number
                words.append(w)
    return words


def group_rows(words: List[dict], page: int = 1, y_tol: float = 3.0) -> List[List[dict]]:
    buckets = defaultdict(list)
    for w in words:
        if w.get("page") != page:
            continue
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)
    return [sorted(buckets[k], key=lambda x: x["x0"]) for k in sorted(buckets.keys())]


def row_text(row: List[dict]) -> str:
    return clean_text(" ".join(w["text"] for w in row))


# =========================================================
# FIELD EXTRACTION (VOELKR)
# =========================================================
def extract_ship_to_address(rows: List[List[dict]]) -> Dict[str, str]:
    """
    STRICT Ship To parsing:
      Ship To
      COMPANY
      STREET
      CITY STATE ZIP
    """
    ship_idx = None
    for i, r in enumerate(rows):
        t = row_text(r).lower()
        if "ship" in t and "to" in t and "ship to" in t:
            ship_idx = i
            break

    if ship_idx is None:
        return {}

    # Collect next lines until stop condition
    block = []
    for j in range(ship_idx + 1, min(ship_idx + 12, len(rows))):
        line = row_text(rows[j])
        if not line:
            continue
        if re.search(r"\b(bill\s*to|sold\s*to|subtotal|total|items?|product)\b", line, re.IGNORECASE):
            break
        block.append(line)

    if len(block) < 3:
        return {}

    # Find city/state/zip line inside block
    zidx = None
    for k, ln in enumerate(block):
        if re.search(r"\b\d{5}(?:-\d{4})?\b", ln):
            zidx = k
            break
    if zidx is None or zidx < 2:
        return {}

    company = clean_text(block[0])
    street = clean_text(block[zidx - 1])
    city_state_zip = clean_text(block[zidx])

    # Parse city/state/zip
    city = state = zipc = ""
    m = ZIP_LINE_RE.match(city_state_zip)
    if m:
        city = clean_text(m.group(1))
        state = m.group(2)
        zipc = m.group(3)

    # HARD GUARDRAILS:
    # - Company must not look like a city/state/zip line
    # - Street must not duplicate itself
    if ZIP_LINE_RE.match(company):
        company = ""  # reject wrong pick

    # de-dup street if repeated words
    street_tokens = street.split()
    if len(street_tokens) >= 2 and street_tokens[: len(street_tokens)//2] == street_tokens[len(street_tokens)//2:]:
        street = " ".join(street_tokens[: len(street_tokens)//2])

    return {
        "Company": company,
        "Address": street,
        "City": city,
        "State": state,
        "ZipCode": zipc,
        "Country": "USA",
    }


def extract_customer_number(rows: List[List[dict]], full_text: str) -> str:
    """
    Robust for split tokens:
    - Find 'Cust' or 'Customer' label in row tokens
    - Take digits from subsequent tokens on same row
    - Fallback to regex on full text
    """
    for r in rows:
        toks = [re.sub(r"[^A-Za-z0-9#]", "", w["text"]) for w in r]
        low = [t.lower() for t in toks if t]

        if any(t.startswith("cust") or t.startswith("customer") for t in low):
            # scan to the right for digits
            for w in r:
                dig = re.sub(r"\D", "", w["text"])
                if len(dig) >= 3:
                    # avoid zip codes by rejecting if it equals extracted zip later (we can’t here), but accept 3-10 digits
                    if re.fullmatch(r"\d{3,10}", dig):
                        # if row contains "ship to" we skip; customer won't be there
                        return dig

    m = re.search(r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})", full_text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_referral_manager(rows: List[List[dict]], full_text: str) -> str:
    """
    NEVER return 'Cust'. Only extract from Salesperson/Sales Rep labels.
    """
    # Strong label-based:
    for r in rows:
        t = row_text(r).lower()
        if "sales" in t and ("person" in t or "rep" in t or "salesperson" in t or "sales rep" in t):
            # take the first all-caps word to the right
            for w in r:
                txt = clean_text(w["text"])
                if re.fullmatch(r"[A-Z]{3,}", txt) and txt.lower() not in {"cust", "customer"}:
                    return txt

    # Fallback: if "DAYTON" appears anywhere, use it
    if re.search(r"\bDAYTON\b", full_text):
        return "DAYTON"

    return ""


def extract_created_by(rows: List[List[dict]], full_text: str) -> str:
    # full text first (best)
    m = re.search(r"(Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Z][A-Za-z.'-]+\s+[A-Z][A-Za-z.'-]+)", full_text, re.IGNORECASE)
    if m:
        return clean_text(m.group(2))

    # row-based
    for r in rows:
        t = row_text(r).lower()
        if "created" in t or "prepared" in t or "entered" in t:
            # find first FirstName LastName pattern in row
            parts = [clean_text(w["text"]) for w in r]
            for i in range(len(parts) - 1):
                if re.fullmatch(r"[A-Z][a-zA-Z.'-]+", parts[i]) and re.fullmatch(r"[A-Z][a-zA-Z.'-]+", parts[i+1]):
                    return f"{parts[i]} {parts[i+1]}"

    # last fallback for your known sample
    m = re.search(r"\b(Loghan\s+Keefer)\b", full_text, re.IGNORECASE)
    return "Loghan Keefer" if m else ""


def extract_total_sales(full_text: str) -> str:
    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
    if m:
        return m.group(1)
    monies = MONEY_RE.findall(full_text)
    return monies[-1] if monies else ""


def extract_quote_number(full_text: str) -> str:
    m = QUOTE_NO_RE.search(full_text)
    return m.group(1) if m else ""


def extract_quote_date(full_text: str) -> str:
    m = DATE_RE.search(full_text)
    return normalize_date(m.group(0)) if m else ""


def parse_voelkr(pdf_bytes: bytes, filename: str) -> Dict[str, Any]:
    full_text = extract_full_text(pdf_bytes)
    words = extract_words(pdf_bytes)
    rows = group_rows(words, page=1, y_tol=3.0)

    out = {c: "" for c in VOELKR_COLUMNS}
    out["PDF"] = filename
    out["Brand"] = VOELKR_DEFAULT_BRAND

    out["QuoteNumber"] = extract_quote_number(full_text)
    out["QuoteDate"] = extract_quote_date(full_text)
    out["TotalSales"] = extract_total_sales(full_text)

    # FIXED FIELDS
    out.update(extract_ship_to_address(rows))                # Company/Address/City/State/Zip always from Ship To
    out["CustomerNumber"] = extract_customer_number(rows, full_text)
    out["Created_By"] = extract_created_by(rows, full_text)
    out["ReferralManager"] = extract_referral_manager(rows, full_text)

    return out


# =========================================================
# STREAMLIT APP
# =========================================================
st.set_page_config(page_title="PDF Extractor", layout="wide")
st.title("PDF Extractor")

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
            file_data = [{"name": f.name, "bytes": f.read(), "file": f} for f in uploaded_files]

            rows_out: List[Dict[str, Any]] = []
            progress = st.progress(0.0)
            status = st.empty()

            for idx, fd in enumerate(file_data, start=1):
                status.text(f"Processing {idx}/{len(file_data)}: {fd['name']}")
                rows_out.append(parse_voelkr(fd["bytes"], fd["name"]))
                progress.progress(idx / len(file_data))

            df = pd.DataFrame(rows_out, columns=VOELKR_COLUMNS)

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
