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


def group_words_into_lines(words: List[dict], y_tol=3.0) -> List[str]:
    buckets = defaultdict(list)
    for w in words:
        if w.get("page_num") != 1:
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
    # Prefer label-based
    for ln in lines:
        if "date" in ln.lower():
            m = DATE_RE.search(ln)
            if m:
                return normalize_date(m.group(1))

    # Near quote number lines
    if quote_number:
        for i, ln in enumerate(lines):
            if quote_number in ln:
                for j in range(i, min(i + 10, len(lines))):
                    m = DATE_RE.search(lines[j])
                    if m:
                        return normalize_date(m.group(1))

    # Fallback: first date in full text
    m = DATE_RE.search(full_text)
    return normalize_date(m.group(1)) if m else ""


def extract_address_block_from_lines(lines: List[str]) -> Dict[str, str]:
    out: Dict[str, str] = {}

    for i, ln in enumerate(lines):
        if not ZIP_RE.search(ln):
            continue

        zip_code = ZIP_RE.search(ln).group(1)
        tokens = ln.split()

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

        out["ZipCode"] = zip_code
        if state:
            out["State"] = state
        if city:
            out["City"] = city

        # Street line above
        if i - 1 >= 0:
            out["Address"] = lines[i - 1].strip()

        # Company line above street
        if i - 2 >= 0:
            cand = lines[i - 2].strip()
            if cand and not re.search(r"(date|quote|customer|phone|fax|ship|bill|sold\s*to)", cand, re.IGNORECASE):
                out["Company"] = cand

        # Country
        if i + 1 < len(lines) and USA_RE.search(lines[i + 1]):
            out["Country"] = "USA"
        elif USA_RE.search(ln):
            out["Country"] = "USA"
        else:
            for k in range(i, min(i + 5, len(lines))):
                if USA_RE.search(lines[k]):
                    out["Country"] = "USA"
                    break

        if out.get("Country") != "USA":
            out["Country"] = "USA"

        if out.get("ZipCode") and out.get("Address"):
            break

    return out


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

    label_re = re.compile(r"^(cust|customer)$", re.IGNORECASE)
    label_compact_re = re.compile(r"^(cust|customer)\#?$", re.IGNORECASE)
    num_re = re.compile(r"^\d{3,10}$")

    for idx, rk in enumerate(sorted_row_keys):
        row = sorted(rows[rk], key=lambda x: x["x0"])

        for w in row:
            t = (w.get("text") or "").strip()
            if not t:
                continue

            if label_re.match(t) or label_compact_re.match(t) or t.lower().startswith("cust"):
                label_x = w["x0"]

                # same row right side
                right_words = sorted([ww for ww in row if ww["x0"] > label_x], key=lambda x: x["x0"])
                for ww in right_words[:15]:
                    val = (ww.get("text") or "").strip().replace(",", "")
                    val = re.sub(r"[^\d]", "", val)
                    if num_re.match(val):
                        return val

                # next rows (1-2)
                for down in (1, 2):
                    if idx + down >= len(sorted_row_keys):
                        break
                    row2 = sorted(rows[sorted_row_keys[idx + down]], key=lambda x: x["x0"])
                    nearby = sorted([ww for ww in row2 if ww["x0"] >= label_x - 8], key=lambda x: x["x0"])
                    for ww in nearby[:15]:
                        val = (ww.get("text") or "").strip().replace(",", "")
                        val = re.sub(r"[^\d]", "", val)
                        if num_re.match(val):
                            return val

    return ""


def extract_customer_number(lines: List[str], full_text: str, words: List[dict]) -> str:
    # 1) word-coordinate method
    v = extract_customer_number_from_words(words)
    if v:
        return v

    # 2) line-based regex
    for ln in lines:
        if re.search(r"\b(cust|customer)\b", ln, re.IGNORECASE):
            m = re.search(
                r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})",
                ln,
                re.IGNORECASE,
            )
            if m:
                return m.group(1)

    # 3) full-text regex
    m = re.search(
        r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})",
        full_text,
        re.IGNORECASE,
    )
    return m.group(1) if m else ""


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    words = extract_words_all_pages(pdf_bytes)
    lines = group_words_into_lines(words)
    full_text = extract_full_text(pdf_bytes)

    out: Dict[str, str] = {}

    # QuoteNumber
    m = QUOTE_NO_RE.search(full_text)
    if m:
        out["QuoteNumber"] = m.group(1).strip()

    # QuoteDate
    out["QuoteDate"] = extract_quote_date(lines, full_text, out.get("QuoteNumber", ""))

    # CustomerNumber (fixed)
    out["CustomerNumber"] = extract_customer_number(lines, full_text, words)

    # ReferralManager (sample: DAYTON)
    if any(re.search(r"\bDAYTON\b", ln) for ln in lines) or re.search(r"\bDAYTON\b", full_text):
        out["ReferralManager"] = "DAYTON"

    # TotalSales
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

    # Address block
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
st.title("PDF Extractor – Voelkr")

st.markdown(
    """
- **Voelkr** extracts header fields (includes Cust # / CustomerNumber).
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

# Show results if present
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
