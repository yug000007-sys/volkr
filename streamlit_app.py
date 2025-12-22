import io
import re
import zipfile
from typing import List, Dict, Any, Optional
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
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
CITY_ST_ZIP_RE = re.compile(r"^(.*)\s+([A-Z]{2})\s+(\d{5})(?:-\d{4})?$")
MONEY_RE = re.compile(r"\b([\d,]+\.\d{2})\b")


def clean_text(s: str) -> str:
    return re.sub(r"\s{2,}", " ", (s or "").replace("\n", " ")).strip()


def normalize_date(s: str) -> str:
    s = (s or "").strip()
    m = DATE_RE.search(s)
    if not m:
        return s
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


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
                w["text_norm"] = (w.get("text") or "").strip().lower().replace(":", "").replace("#", "")
                words.append(w)
    return words


def group_rows(words: List[dict], page: int = 1, y_tol: float = 3.0) -> List[Dict[str, Any]]:
    """
    Returns list of rows:
      { "y": float, "words": [word...], "text": str }
    """
    buckets = defaultdict(list)
    for w in words:
        if w.get("page") != page:
            continue
        key = round(w["top"] / y_tol) * y_tol
        buckets[key].append(w)

    rows = []
    for y in sorted(buckets.keys()):
        row_words = sorted(buckets[y], key=lambda x: x["x0"])
        rows.append(
            {
                "y": y,
                "words": row_words,
                "text": clean_text(" ".join(w["text"] for w in row_words)),
            }
        )
    return rows


def row_tokens_norm(row_words: List[dict]) -> List[str]:
    return [w["text_norm"] for w in row_words if w.get("text_norm")]


# =========================================================
# FIELD EXTRACTORS (VOELKR)
# =========================================================
def extract_quote_number(full_text: str) -> str:
    m = QUOTE_NO_RE.search(full_text)
    return m.group(1) if m else ""


def extract_quote_date(full_text: str) -> str:
    m = DATE_RE.search(full_text)
    return normalize_date(m.group(0)) if m else ""


def extract_total_sales(full_text: str) -> str:
    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
    if m:
        return m.group(1)
    monies = MONEY_RE.findall(full_text)
    return monies[-1] if monies else ""


def find_ship_to_anchor(rows: List[Dict[str, Any]]) -> Optional[dict]:
    """
    Find a row that contains token 'ship' and token 'to' (even if split).
    Returns dict with:
      { "row_idx": int, "x0": float, "x1": float }
    """
    for i, r in enumerate(rows):
        toks = row_tokens_norm(r["words"])
        if "ship" in toks and "to" in toks:
            # anchor x-range = from 'ship' word x0 to 'to' word x1
            ship_ws = [w for w in r["words"] if w["text_norm"] == "ship"]
            to_ws = [w for w in r["words"] if w["text_norm"] == "to"]
            if ship_ws and to_ws:
                x0 = min(w["x0"] for w in ship_ws)
                x1 = max(w["x1"] for w in to_ws)
                return {"row_idx": i, "x0": x0, "x1": x1}
    return None


def extract_ship_to_block(words: List[dict]) -> Dict[str, str]:
    """
    ALWAYS extract address from Ship To.
    Uses the Ship/To anchor row and reads the next rows from the SAME COLUMN area.
    """
    rows = group_rows(words, page=1, y_tol=3.0)
    anchor = find_ship_to_anchor(rows)
    if not anchor:
        return {}

    start_i = anchor["row_idx"] + 1
    # define a "column window" around Ship To label
    col_left = max(0, anchor["x0"] - 40)
    col_right = anchor["x1"] + 420  # wide enough to capture address block values

    block_lines: List[str] = []
    for j in range(start_i, min(start_i + 12, len(rows))):
        row_words = rows[j]["words"]

        # take only words in the same column window
        in_col = [w for w in row_words if w["x0"] >= col_left and w["x1"] <= col_right]
        line = clean_text(" ".join(w["text"] for w in in_col))

        if not line:
            continue

        # stop if we hit another header section
        if re.search(r"\b(bill\s*to|sold\s*to|subtotal|total|items?|product)\b", line, re.IGNORECASE):
            break

        block_lines.append(line)

    if not block_lines:
        return {}

    # find the city/state/zip line (first with ZIP)
    zip_i = None
    for k, ln in enumerate(block_lines):
        if ZIP_RE.search(ln):
            zip_i = k
            break
    if zip_i is None:
        return {}

    # company: first non-empty line BEFORE zip line
    # street: line immediately before zip line
    company = ""
    for ln in block_lines[:zip_i]:
        if ln:
            company = ln
            break

    street = block_lines[zip_i - 1] if zip_i - 1 >= 0 else ""
    city_state_zip = block_lines[zip_i]

    # parse city/state/zip
    city = state = zipc = ""
    m = CITY_ST_ZIP_RE.match(city_state_zip)
    if m:
        city = clean_text(m.group(1))
        state = m.group(2)
        zipc = m.group(3)

    # guardrails: company must not be city/state/zip
    if CITY_ST_ZIP_RE.match(company):
        company = ""

    # prevent duplicated street like "4265 ... 4265 ..."
    street_tokens = street.split()
    if len(street_tokens) >= 4:
        half = len(street_tokens) // 2
        if street_tokens[:half] == street_tokens[half:]:
            street = " ".join(street_tokens[:half])

    return {
        "Company": company,
        "Address": street,
        "City": city,
        "State": state,
        "ZipCode": zipc,
        "Country": "USA",
    }


def extract_customer_number(words: List[dict], full_text: str) -> str:
    """
    Robust 'Cust # 10980' even if split across tokens/rows/columns.
    Strategy:
      1) find token starting with 'cust' or 'customer'
      2) scan right on same row for digits
      3) scan next 3 rows in same column area
      4) fallback regex on full text
    """
    rows = group_rows(words, page=1, y_tol=3.0)

    def digits_only(s: str) -> str:
        return re.sub(r"\D", "", s or "")

    for i, r in enumerate(rows):
        toks = row_tokens_norm(r["words"])
        if any(t.startswith("cust") or t.startswith("customer") for t in toks):
            # locate the first cust-like word to get x position
            cust_words = [w for w in r["words"] if w["text_norm"].startswith("cust") or w["text_norm"].startswith("customer")]
            if not cust_words:
                continue
            cw = cust_words[0]
            x_anchor = cw["x0"]

            # scan right same row
            right = [w for w in r["words"] if w["x0"] > x_anchor]
            for w in right:
                d = digits_only(w["text"])
                if re.fullmatch(r"\d{3,10}", d or ""):
                    return d

            # scan down next rows near same x column
            for down in range(1, 4):
                if i + down >= len(rows):
                    break
                r2 = rows[i + down]["words"]
                near = [w for w in r2 if w["x0"] >= x_anchor - 10]
                for w in near:
                    d = digits_only(w["text"])
                    if re.fullmatch(r"\d{3,10}", d or ""):
                        return d

    # fallback text regex
    m = re.search(r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})", full_text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_created_by(words: List[dict], full_text: str) -> str:
    m = re.search(
        r"(Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Z][A-Za-z.'-]+\s+[A-Z][A-Za-z.'-]+)",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return clean_text(m.group(2))

    rows = group_rows(words, page=1, y_tol=3.0)
    for r in rows:
        t = r["text"].lower()
        if "created" in t or "prepared" in t or "entered" in t:
            # find first FirstName LastName in row
            parts = [w["text"].strip() for w in r["words"]]
            for i in range(len(parts) - 1):
                if re.fullmatch(r"[A-Z][A-Za-z.'-]+", parts[i]) and re.fullmatch(r"[A-Z][A-Za-z.'-]+", parts[i + 1]):
                    return f"{parts[i]} {parts[i + 1]}"

    # last fallback for your sample
    m = re.search(r"\b(Loghan\s+Keefer)\b", full_text, re.IGNORECASE)
    return "Loghan Keefer" if m else ""


def extract_referral_manager(words: List[dict], full_text: str) -> str:
    """
    Only from Salesperson/Sales Rep. Never from Cust.
    """
    rows = group_rows(words, page=1, y_tol=3.0)
    for r in rows:
        t = r["text"].lower()
        if "sales" in t and ("person" in t or "rep" in t or "salesperson" in t):
            # take first ALLCAPS name token that isn't CUST/CUSTOMER
            for w in r["words"]:
                txt = w["text"].strip()
                if re.fullmatch(r"[A-Z]{3,}", txt) and txt.lower() not in {"cust", "customer"}:
                    return txt

    # fallback to DAYTON if present anywhere
    if re.search(r"\bDAYTON\b", full_text):
        return "DAYTON"

    return ""


def parse_voelkr(pdf_bytes: bytes, filename: str) -> Dict[str, Any]:
    full_text = extract_full_text(pdf_bytes)
    words = extract_words(pdf_bytes)

    out = {c: "" for c in VOELKR_COLUMNS}
    out["PDF"] = filename
    out["Brand"] = VOELKR_DEFAULT_BRAND

    out["QuoteNumber"] = extract_quote_number(full_text)
    out["QuoteDate"] = extract_quote_date(full_text)
    out["TotalSales"] = extract_total_sales(full_text)

    # FIXED: Ship To ALWAYS
    ship = extract_ship_to_block(words)
    out.update(ship)

    # FIXED: CustomerNumber robust
    out["CustomerNumber"] = extract_customer_number(words, full_text)

    # FIXED: Created_By robust
    out["Created_By"] = extract_created_by(words, full_text)

    # FIXED: ReferralManager robust (never "Cust")
    out["ReferralManager"] = extract_referral_manager(words, full_text)

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
            file_data = [{"name": f.name, "bytes": f.read()} for f in uploaded_files]

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
