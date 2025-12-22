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
    return m.group(0) if m else s.strip()


def clean_company_line(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"^\s*(ship\s*to|ship-to)\s*:?\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^\s*(sold\s*to|bill\s*to)\s*:?\s*", "", s, flags=re.IGNORECASE)
    return s.strip()


# =========================================================
# PDF WORD EXTRACTION
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
# REGEX
# =========================================================
DATE_RE = re.compile(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b")
QUOTE_NO_RE = re.compile(r"\b(\d{2}/\d{6})\b")
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
USA_RE = re.compile(r"\b(USA|United States of America)\b", re.IGNORECASE)
MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


# =========================================================
# VOELKR FIELD EXTRACTORS
# =========================================================
def extract_quote_date(lines: List[str], full_text: str, quote_number: str) -> str:
    for ln in lines:
        if "date" in ln.lower():
            m = DATE_RE.search(ln)
            if m:
                return normalize_date(m.group(1))

    if quote_number:
        for i, ln in enumerate(lines):
            if quote_number in ln:
                for j in range(i, min(i + 12, len(lines))):
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

    for idx, rk in enumerate(sorted_row_keys):
        row = sorted(rows[rk], key=lambda x: x["x0"])
        for w in row:
            t = (w.get("text") or "").strip()
            if not t:
                continue
            if t.lower().startswith("cust") or t.lower().startswith("customer"):
                label_x = w["x0"]

                right = sorted([ww for ww in row if ww["x0"] > label_x], key=lambda x: x["x0"])
                for ww in right[:25]:
                    val = re.sub(r"[^\d]", "", (ww.get("text") or ""))
                    if num_re.match(val):
                        return val

                for down in (1, 2, 3):
                    if idx + down >= len(sorted_row_keys):
                        break
                    row2 = sorted(rows[sorted_row_keys[idx + down]], key=lambda x: x["x0"])
                    near = sorted([ww for ww in row2 if ww["x0"] >= label_x - 10], key=lambda x: x["x0"])
                    for ww in near[:25]:
                        val = re.sub(r"[^\d]", "", (ww.get("text") or ""))
                        if num_re.match(val):
                            return val

    return ""


def extract_created_by_from_words(words: List[dict]) -> str:
    """
    Robust Created_By:
      - finds 'Created' or 'Prepared' label in word coords
      - captures name tokens to the right on same row
      - if 'By' is separate token, handles that too
    """
    if not words:
        return ""

    w1 = [w for w in words if w.get("page_num") == 1]

    def row_key(w, y_tol=3.0):
        return round(w["top"] / y_tol) * y_tol

    rows = defaultdict(list)
    for w in w1:
        rows[row_key(w)].append(w)

    sorted_keys = sorted(rows.keys())

    def looks_like_name(tok: str) -> bool:
        tok = (tok or "").strip()
        if not tok:
            return False
        if re.search(r"\d", tok):
            return False
        if tok.lower() in {"date", "quote", "customer", "ship", "bill", "sold", "to", "total", "subtotal"}:
            return False
        return bool(re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", tok))

    for rk in sorted_keys:
        row = sorted(rows[rk], key=lambda x: x["x0"])
        texts = [(w["x0"], (w.get("text") or "").strip()) for w in row]
        for i, (x0, t) in enumerate(texts):
            tl = t.lower()
            if tl in {"created", "prepared", "entered"} or tl.startswith("created") or tl.startswith("prepared"):
                # skip optional "by"
                j = i + 1
                if j < len(texts) and texts[j][1].lower() == "by":
                    j += 1

                name_parts = []
                for k in range(j, min(j + 6, len(texts))):
                    tok = texts[k][1]
                    if looks_like_name(tok):
                        name_parts.append(tok)
                    else:
                        break

                val = " ".join(name_parts).strip()
                if len(val.split()) >= 2:
                    return val

    # fallback: search full-text label patterns (sometimes works even if word-splitting fails)
    return ""


def extract_ship_to_address_block(lines: List[str]) -> Dict[str, str]:
    """
    Extract Ship To block ONLY (your requirement).
    We:
      1) find 'Ship To' marker (even if embedded in a longer line)
      2) scan forward up to 20 lines until we find a line with ZIP (city/state/zip)
      3) street = line immediately above the ZIP line
      4) company = first non-label line between Ship To and street (usually the first line)
    """
    out: Dict[str, str] = {}
    ship_idx = -1

    for i, ln in enumerate(lines):
        if re.search(r"\bship\s*to\b", ln, re.IGNORECASE):
            ship_idx = i
            break

    if ship_idx == -1:
        return out

    # collect candidate block lines after Ship To
    cand = []
    for j in range(ship_idx + 1, min(ship_idx + 25, len(lines))):
        t = lines[j].strip()
        if not t:
            continue
        # stop if we hit totals/items section
        if re.search(r"\b(subtotal|total|grand\s*total|items?|product)\b", t, re.IGNORECASE):
            break
        cand.append(t)

    if not cand:
        return out

    # Find ZIP line in cand
    zip_k = -1
    for k, ln in enumerate(cand):
        if ZIP_RE.search(ln):
            zip_k = k
            break
    if zip_k == -1:
        return out

    city_state_zip = cand[zip_k]
    zip_code = ZIP_RE.search(city_state_zip).group(1)

    tokens = city_state_zip.split()
    state = ""
    if zip_code in tokens:
        zidx = tokens.index(zip_code)
        if zidx >= 1 and re.fullmatch(r"[A-Z]{2}", tokens[zidx - 1]):
            state = tokens[zidx - 1]

    city = ""
    if state and state in tokens:
        sidx = tokens.index(state)
        if sidx >= 1:
            city = " ".join(tokens[:sidx]).strip()

    # Street line is previous line in cand
    street = cand[zip_k - 1] if zip_k - 1 >= 0 else ""

    # Company: best candidate is the first non-label line in cand (often cand[0])
    company = ""
    for ln in cand[:zip_k]:
        if re.search(r"\b(bill\s*to|sold\s*to|ship\s*to)\b", ln, re.IGNORECASE):
            continue
        if ZIP_RE.search(ln):
            continue
        # avoid picking street as company
        if ln == street:
            continue
        company = ln
        break

    company = clean_company_line(company)

    if company:
        out["Company"] = company
    if street:
        out["Address"] = street.strip()
    if city:
        out["City"] = city
    if state:
        out["State"] = state
    if zip_code:
        out["ZipCode"] = zip_code

    # Country: check nearby lines or default USA
    out["Country"] = "USA"
    for ln in cand[zip_k: min(zip_k + 4, len(cand))]:
        if USA_RE.search(ln):
            out["Country"] = "USA"
            break

    return out


def extract_total_sales(lines: List[str], full_text: str) -> str:
    for ln in lines:
        if re.search(r"\b(total|grand total|total sales)\b", ln, re.IGNORECASE):
            mm = MONEY_RE.search(ln)
            if mm:
                return normalize_money(mm.group(0))

    m = re.search(
        r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        return normalize_money(m.group(1))

    monies = MONEY_RE.findall(full_text)
    return normalize_money(monies[-1]) if monies else ""


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

    # ReferralManager (sample)
    if any(re.search(r"\bDAYTON\b", ln) for ln in lines) or re.search(r"\bDAYTON\b", full_text):
        out["ReferralManager"] = "DAYTON"

    # TotalSales
    out["TotalSales"] = extract_total_sales(lines, full_text)

    # Created_By (word-based; fixes split-label cases)
    out["Created_By"] = extract_created_by_from_words(words)

    # SHIP TO address (forces correct company/address/city/state/zip)
    out.update(extract_ship_to_address_block(lines))

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
- **Voelkr** extracts header fields using **Ship To** for Company/Address/City/State/Zip.
- **Created_By** extracted using **word-coordinate** logic (handles split labels).
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
