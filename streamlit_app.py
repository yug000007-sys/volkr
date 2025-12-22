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
    for i, r in enumerate(rows):
        toks = row_tokens_norm(r["words"])
        if "ship" in toks and "to" in toks:
            ship_ws = [w for w in r["words"] if w["text_norm"] == "ship"]
            to_ws = [w for w in r["words"] if w["text_norm"] == "to"]
            if ship_ws and to_ws:
                x0 = min(w["x0"] for w in ship_ws)
                x1 = max(w["x1"] for w in to_ws)
                return {"row_idx": i, "x0": x0, "x1": x1}
    return None


def _looks_like_city_state_zip(s: str) -> bool:
    return bool(CITY_ST_ZIP_RE.match(clean_text(s)))


def _looks_like_street(s: str) -> bool:
    s = clean_text(s)
    # typical street starts with number
    return bool(re.match(r"^\d{2,6}\s+", s))


def extract_ship_to_block(words: List[dict]) -> Dict[str, str]:
    rows = group_rows(words, page=1, y_tol=3.0)
    anchor = find_ship_to_anchor(rows)
    if not anchor:
        return {}

    start_i = anchor["row_idx"] + 1
    col_left = max(0, anchor["x0"] - 40)
    col_right = anchor["x1"] + 420

    block_lines: List[str] = []
    for j in range(start_i, min(start_i + 14, len(rows))):
        row_words = rows[j]["words"]
        in_col = [w for w in row_words if w["x0"] >= col_left and w["x1"] <= col_right]
        line = clean_text(" ".join(w["text"] for w in in_col))
        if not line:
            continue
        if re.search(r"\b(bill\s*to|sold\s*to|subtotal|total|items?|product)\b", line, re.IGNORECASE):
            break
        block_lines.append(line)

    if not block_lines:
        return {}

    # city/state/zip line = first line containing ZIP
    zip_i = None
    for k, ln in enumerate(block_lines):
        if ZIP_RE.search(ln):
            zip_i = k
            break
    if zip_i is None:
        return {}

    city_state_zip = block_lines[zip_i]
    m = CITY_ST_ZIP_RE.match(city_state_zip)
    city = state = zipc = ""
    if m:
        city = clean_text(m.group(1))
        state = m.group(2)
        zipc = m.group(3)

    # street line = nearest line ABOVE city_state_zip that looks like street
    street = ""
    for k in range(zip_i - 1, -1, -1):
        if _looks_like_street(block_lines[k]):
            street = block_lines[k]
            break

    # company line = nearest line ABOVE street that is NOT street and NOT city/state/zip
    company = ""
    if street:
        street_i = block_lines.index(street)
        for k in range(street_i - 1, -1, -1):
            cand = block_lines[k]
            if not cand:
                continue
            if _looks_like_city_state_zip(cand):
                continue
            if _looks_like_street(cand):
                continue
            # reject if it contains only repeated city/state/zip fragments
            company = cand
            break
    else:
        # fallback: first line that is not city/state/zip
        for cand in block_lines[:zip_i]:
            if cand and not _looks_like_city_state_zip(cand):
                company = cand
                break

    # clean duplication in street
    street_tokens = street.split()
    if len(street_tokens) >= 4:
        half = len(street_tokens) // 2
        if street_tokens[:half] == street_tokens[half:]:
            street = " ".join(street_tokens[:half])

    # guardrails: company must not be city/state/zip or a street
    if _looks_like_city_state_zip(company) or _looks_like_street(company):
        company = ""

    return {
        "Company": company,
        "Address": street,
        "City": city,
        "State": state,
        "ZipCode": zipc,
        "Country": "USA",
    }


def extract_customer_number(words: List[dict], full_text: str) -> str:
    rows = group_rows(words, page=1, y_tol=3.0)

    def digits_only(s: str) -> str:
        return re.sub(r"\D", "", s or "")

    for i, r in enumerate(rows):
        toks = row_tokens_norm(r["words"])
        if any(t.startswith("cust") or t.startswith("customer") for t in toks):
            cust_words = [w for w in r["words"] if w["text_norm"].startswith("cust") or w["text_norm"].startswith("customer")]
            if not cust_words:
                continue
            cw = cust_words[0]
            x_anchor = cw["x0"]

            right = [w for w in r["words"] if w["x0"] > x_anchor]
            for w in right:
                d = digits_only(w["text"])
                if re.fullmatch(r"\d{3,10}", d or ""):
                    return d

            for down in range(1, 4):
                if i + down >= len(rows):
                    break
                r2 = rows[i + down]["words"]
                near = [w for w in r2 if w["x0"] >= x_anchor - 10]
                for w in near:
                    d = digits_only(w["text"])
                    if re.fullmatch(r"\d{3,10}", d or ""):
                        return d

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
    m = re.search(r"\b(Loghan\s+Keefer)\b", full_text, re.IGNORECASE)
    return "Loghan Keefer" if m else ""


def extract_referral_manager(words: List[dict], full_text: str) -> str:
    rows = group_rows(words, page=1, y_tol=3.0)
    for r in rows:
        t = r["text"].lower()
        if "sales" in t and ("person" in t or "rep" in t or "salesperson" in t):
            for w in r["words"]:
                txt = w["text"].strip()
                if re.fullmatch(r"[A-Z]{3,}", txt) and txt.lower() not in {"cust", "customer"}:
                    return txt
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

    ship = extract_ship_to_block(words)
    out.update(ship)

    out["CustomerNumber"] = extract_customer_number(words, full_text)
    out["Created_By"] = extract_created_by(words, full_text)
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
