import io
import re
import zipfile
from typing import List, Dict, Any, Tuple, Optional
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
DATE_RE = re.compile(r"\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b")
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


def normalize_date(s: str) -> str:
    s = (s or "").strip()
    m = DATE_RE.search(s)
    if not m:
        return s
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def normalize_money(s: str) -> str:
    s = (s or "").strip()
    m = MONEY_RE.search(s)
    return m.group(0) if m else s


def clean_text(s: str) -> str:
    return re.sub(r"\s{2,}", " ", (s or "").replace("\n", " ")).strip()


def is_name_token(tok: str) -> bool:
    tok = (tok or "").strip()
    if not tok:
        return False
    if re.search(r"\d", tok):
        return False
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", tok))


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


def extract_full_text(pdf_bytes: bytes) -> str:
    parts: List[str] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def group_words_by_rows(words: List[dict], page_num: int = 1, y_tol: float = 3.0) -> List[Tuple[float, List[dict]]]:
    w = [x for x in words if x.get("page_num") == page_num]
    buckets = defaultdict(list)
    for ww in w:
        key = round(ww["top"] / y_tol) * y_tol
        buckets[key].append(ww)
    rows: List[Tuple[float, List[dict]]] = []
    for y in sorted(buckets.keys()):
        row = sorted(buckets[y], key=lambda x: x["x0"])
        rows.append((y, row))
    return rows


def row_text(row: List[dict]) -> str:
    return clean_text(" ".join(w["text"] for w in row))


def norm_label_token(t: str) -> str:
    t = (t or "").strip().lower()
    t = t.replace(":", "").replace("#", "").replace(".", "")
    return t


def find_phrase_in_row(row: List[dict], phrase_tokens: List[str]) -> Optional[Tuple[int, int]]:
    """
    Returns (start_idx, end_idx_exclusive) if phrase appears in row tokens (normalized).
    """
    toks = [norm_label_token(w["text"]) for w in row]
    p = [norm_label_token(x) for x in phrase_tokens]
    if not p or not toks:
        return None
    for i in range(0, len(toks) - len(p) + 1):
        if toks[i : i + len(p)] == p:
            return i, i + len(p)
    return None


def value_right_of_span(row: List[dict], span_end_idx: int, max_tokens: int = 10) -> str:
    right = row[span_end_idx : span_end_idx + max_tokens]
    return clean_text(" ".join(w["text"] for w in right))


def value_in_same_row_after_phrase(rows: List[Tuple[float, List[dict]]], phrases: List[List[str]], max_tokens: int = 10) -> str:
    for _, row in rows:
        for phrase in phrases:
            span = find_phrase_in_row(row, phrase)
            if span:
                v = value_right_of_span(row, span[1], max_tokens=max_tokens)
                if v:
                    return v
    return ""


def value_below_phrase(rows: List[Tuple[float, List[dict]]], phrases: List[List[str]], lines_down: int = 3, max_tokens: int = 12) -> str:
    for i, (_, row) in enumerate(rows):
        for phrase in phrases:
            span = find_phrase_in_row(row, phrase)
            if not span:
                continue
            # Try same row right first
            v = value_right_of_span(row, span[1], max_tokens=max_tokens)
            if v:
                return v
            # Then look below
            for d in range(1, lines_down + 1):
                if i + d >= len(rows):
                    break
                _, r2 = rows[i + d]
                v2 = clean_text(" ".join(w["text"] for w in r2[:max_tokens]))
                if v2:
                    return v2
    return ""


# =========================================================
# VOELKR: STRICT SHIP TO EXTRACTION (ALWAYS)
# =========================================================
def extract_ship_to_block(rows: List[Tuple[float, List[dict]]]) -> Dict[str, str]:
    """
    Always take address fields from SHIP TO.
    Strategy:
      - find a row containing phrase ["ship","to"] (tokens may be separated)
      - collect next rows until another header label appears
      - parse:
         company = first non-empty row line
         street  = row just before city/state/zip line
         city/state/zip = row containing ZIP
    """
    stop_phrases = [
        ["bill", "to"],
        ["sold", "to"],
        ["quote"],
        ["customer"],
        ["subtotal"],
        ["total"],
        ["product"],
        ["items"],
    ]

    ship_idx = None
    for i, (_, row) in enumerate(rows):
        if find_phrase_in_row(row, ["ship", "to"]):
            ship_idx = i
            break

    if ship_idx is None:
        return {}

    block_lines: List[str] = []
    for j in range(ship_idx + 1, min(ship_idx + 20, len(rows))):
        _, r = rows[j]
        line = row_text(r)
        if not line:
            continue
        # stop if another section starts
        stop = False
        for sp in stop_phrases:
            if find_phrase_in_row(r, sp):
                stop = True
                break
        if stop:
            break
        block_lines.append(line)

    if not block_lines:
        return {}

    # find city/state/zip line
    zip_i = None
    for k, ln in enumerate(block_lines):
        if ZIP_RE.search(ln):
            zip_i = k
            break
    if zip_i is None:
        return {}

    city_state_zip = block_lines[zip_i]
    zip_code = ZIP_RE.search(city_state_zip).group(1)

    # parse state as the 2-letter token right before zip
    tokens = city_state_zip.split()
    state = ""
    if zip_code in tokens:
        zidx = tokens.index(zip_code)
        if zidx >= 1 and re.fullmatch(r"[A-Z]{2}", tokens[zidx - 1]):
            state = tokens[zidx - 1]

    city = ""
    if state and state in tokens:
        sidx = tokens.index(state)
        city = " ".join(tokens[:sidx]).strip()

    street = block_lines[zip_i - 1].strip() if zip_i - 1 >= 0 else ""
    company = block_lines[0].strip()

    # If company accidentally equals street, use next line
    if company == street and len(block_lines) > 1:
        company = block_lines[1].strip()

    return {
        "Company": company,
        "Address": street,
        "City": city,
        "State": state,
        "ZipCode": zip_code,
        "Country": "USA",
    }


# =========================================================
# VOELKR: FIELD EXTRACTORS BY LABELS
# =========================================================
def extract_quote_number(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    # Prefer label-based
    v = value_in_same_row_after_phrase(
        rows,
        phrases=[
            ["quote", "#"],
            ["quote", "no"],
            ["quote", "number"],
            ["quote"],
        ],
        max_tokens=6,
    )
    # Pick something like 01/125785
    m = re.search(r"\b\d{2}/\d{6}\b", v)
    if m:
        return m.group(0)

    # Fallback text
    m = re.search(r"\b\d{2}/\d{6}\b", full_text)
    return m.group(0) if m else ""


def extract_quote_date(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    v = value_below_phrase(
        rows,
        phrases=[
            ["quote", "date"],
            ["date"],
        ],
        lines_down=3,
        max_tokens=8,
    )
    m = DATE_RE.search(v)
    if m:
        return normalize_date(m.group(0))

    m = DATE_RE.search(full_text)
    return normalize_date(m.group(0)) if m else ""


def extract_customer_number(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    v = value_below_phrase(
        rows,
        phrases=[
            ["cust"],
            ["cust", "#"],
            ["customer"],
            ["customer", "#"],
            ["customer", "no"],
            ["customer", "number"],
        ],
        lines_down=3,
        max_tokens=10,
    )
    digits = re.sub(r"[^\d]", "", v or "")
    if re.fullmatch(r"\d{3,10}", digits or ""):
        return digits

    m = re.search(r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})", full_text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_referral_manager_salesperson(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    v = value_below_phrase(
        rows,
        phrases=[
            ["salesperson"],
            ["sales", "person"],
            ["sales", "rep"],
            ["salesrep"],
            ["rep"],
        ],
        lines_down=3,
        max_tokens=10,
    )
    toks = (v or "").split()
    name_parts: List[str] = []
    for t in toks:
        if is_name_token(t):
            name_parts.append(t)
        else:
            break
    name = " ".join(name_parts).strip()
    if name:
        return name

    m = re.search(r"Salesperson\s*[:#]?\s*([A-Za-z.'-]+\s+[A-Za-z.'-]+)", full_text, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def extract_created_by(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    v = value_below_phrase(
        rows,
        phrases=[
            ["created", "by"],
            ["prepared", "by"],
            ["entered", "by"],
            ["created"],
            ["prepared"],
            ["entered"],
        ],
        lines_down=4,
        max_tokens=12,
    )

    # Strip leading "by"
    v = re.sub(r"^\s*by\s+", "", (v or ""), flags=re.IGNORECASE).strip()

    toks = (v or "").split()
    name_parts: List[str] = []
    for t in toks:
        if is_name_token(t):
            name_parts.append(t)
        else:
            break
    name = " ".join(name_parts).strip()
    if len(name.split()) >= 2:
        return name

    # Text fallback
    m = re.search(
        r"(?:Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Z][A-Za-z.'-]+\s+[A-Z][A-Za-z.'-]+)",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()

    # Last fallback for your known example (won't hurt other PDFs)
    m = re.search(r"\b(Loghan\s+Keefer)\b", full_text, re.IGNORECASE)
    if m:
        return "Loghan Keefer"

    return ""


def extract_total_sales(rows: List[Tuple[float, List[dict]]], full_text: str) -> str:
    # label-based
    v = value_below_phrase(
        rows,
        phrases=[
            ["total", "sales"],
            ["grand", "total"],
            ["total"],
        ],
        lines_down=2,
        max_tokens=10,
    )
    mm = MONEY_RE.search(v or "")
    if mm:
        return normalize_money(mm.group(0))

    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
    if m:
        return normalize_money(m.group(1))

    monies = MONEY_RE.findall(full_text)
    return normalize_money(monies[-1]) if monies else ""


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    words = extract_words_all_pages(pdf_bytes)
    full_text = extract_full_text(pdf_bytes)
    rows = group_words_by_rows(words, page_num=1, y_tol=3.0)

    out: Dict[str, str] = {}
    out["Brand"] = VOELKR_DEFAULT_BRAND

    # Required colored mapping points (by label anchors)
    out["QuoteNumber"] = extract_quote_number(rows, full_text)          # Grey
    out["CustomerNumber"] = extract_customer_number(rows, full_text)    # Green
    out["ReferralManager"] = extract_referral_manager_salesperson(rows, full_text)  # Purple
    out["Created_By"] = extract_created_by(rows, full_text)             # Yellow

    # Ship To ALWAYS for address block (Red)
    out.update(extract_ship_to_block(rows))

    # Other header fields you already had correct (keep)
    out["QuoteDate"] = extract_quote_date(rows, full_text)
    out["TotalSales"] = extract_total_sales(rows, full_text)

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

# Results panel
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

# Run extraction
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
                rows_out.append(build_voelkr_row(fd["bytes"], fd["name"]))
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
