import io
import re
import zipfile
from typing import List, Dict, Any, Optional, Tuple
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


def clean_text(s: str) -> str:
    return re.sub(r"\s{2,}", " ", (s or "").replace("\n", " ")).strip()


def is_nameish(token: str) -> bool:
    token = (token or "").strip()
    if not token:
        return False
    if re.search(r"\d", token):
        return False
    if token.lower() in {"ship", "to", "bill", "sold", "date", "quote", "customer", "total", "subtotal"}:
        return False
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z.'-]*", token))


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
    rows = []
    for y in sorted(buckets.keys()):
        row = sorted(buckets[y], key=lambda x: x["x0"])
        rows.append((y, row))
    return rows


# =========================================================
# GENERIC LABEL -> VALUE (COORDINATE) EXTRACTOR
# =========================================================
def find_label_positions(words: List[dict], label_variants: List[str], page_num: int = 1) -> List[dict]:
    labels = set(v.lower() for v in label_variants)
    out = []
    for w in words:
        if w.get("page_num") != page_num:
            continue
        t = (w.get("text") or "").strip().lower()
        t = t.replace(":", "")
        if t in labels:
            out.append(w)
    return out


def get_text_right_of_label(
    rows: List[Tuple[float, List[dict]]],
    label_word: dict,
    max_tokens: int = 8,
) -> str:
    label_y = round(label_word["top"] / 3.0) * 3.0
    label_x = label_word["x0"]

    # find same row (closest y)
    best_row = None
    best_dy = 9999.0
    for y, row in rows:
        dy = abs(y - label_y)
        if dy < best_dy:
            best_dy = dy
            best_row = row

    if not best_row:
        return ""

    right = [w for w in best_row if w["x0"] > label_x + 3]
    right = sorted(right, key=lambda x: x["x0"])[:max_tokens]
    return clean_text(" ".join(w["text"] for w in right))


def get_text_below_label(
    rows: List[Tuple[float, List[dict]]],
    label_word: dict,
    lines_down: int = 2,
    max_tokens: int = 10,
) -> str:
    label_y = round(label_word["top"] / 3.0) * 3.0
    label_x = label_word["x0"]

    # find row index closest to label_y
    idx = None
    best = 9999.0
    for i, (y, _) in enumerate(rows):
        d = abs(y - label_y)
        if d < best:
            best = d
            idx = i

    if idx is None:
        return ""

    for down in range(1, lines_down + 1):
        if idx + down >= len(rows):
            break
        _, row = rows[idx + down]
        # take tokens around same column or to the right
        cand = [w for w in row if w["x0"] >= label_x - 5]
        cand = sorted(cand, key=lambda x: x["x0"])[:max_tokens]
        txt = clean_text(" ".join(w["text"] for w in cand))
        if txt:
            return txt

    return ""


def extract_value_near_label(
    words: List[dict],
    label_variants: List[str],
    page_num: int = 1,
    prefer_right: bool = True,
    right_max_tokens: int = 8,
    below_lines: int = 3,
) -> str:
    rows = group_words_by_rows(words, page_num=page_num, y_tol=3.0)
    label_words = find_label_positions(words, label_variants, page_num=page_num)
    for lw in label_words:
        if prefer_right:
            v = get_text_right_of_label(rows, lw, max_tokens=right_max_tokens)
            if v:
                return v
            v = get_text_below_label(rows, lw, lines_down=below_lines)
            if v:
                return v
        else:
            v = get_text_below_label(rows, lw, lines_down=below_lines)
            if v:
                return v
            v = get_text_right_of_label(rows, lw, max_tokens=right_max_tokens)
            if v:
                return v
    return ""


# =========================================================
# VOELKR SPECIFIC EXTRACTORS
# =========================================================
QUOTE_NO_RE = re.compile(r"\b(\d{2}/\d{6})\b")
DATE_RE = re.compile(r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\b")
ZIP_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
MONEY_RE = re.compile(r"\b[\d,]+\.\d{2}\b")


def extract_ship_to_block(words: List[dict]) -> Dict[str, str]:
    """
    Strict Ship To:
      Ship To
      COMPANY
      STREET
      CITY STATE ZIP
      (USA)
    Using row index after locating 'Ship' and 'To' (sometimes split).
    """
    rows = group_words_by_rows(words, page_num=1, y_tol=3.0)

    # find a row that contains "Ship" and "To"
    ship_row_idx = None
    for i, (_, row) in enumerate(rows):
        row_text = " ".join(w["text"] for w in row).lower()
        if "ship" in row_text and "to" in row_text:
            # ensure it's actually the label, not in description
            if re.search(r"\bship\b", row_text) and re.search(r"\bto\b", row_text):
                ship_row_idx = i
                break

    if ship_row_idx is None:
        return {}

    # collect next lines after ship_to label
    block_lines = []
    for j in range(ship_row_idx + 1, min(ship_row_idx + 12, len(rows))):
        _, row = rows[j]
        line = clean_text(" ".join(w["text"] for w in row))
        if not line:
            continue
        # stop if we hit other header sections
        if re.search(r"\b(bill\s*to|sold\s*to|subtotal|total|items?|product)\b", line, re.IGNORECASE):
            break
        block_lines.append(line)

    if not block_lines:
        return {}

    # find city/state/zip line by zip
    zip_idx = None
    for k, ln in enumerate(block_lines):
        if ZIP_RE.search(ln):
            zip_idx = k
            break
    if zip_idx is None:
        return {}

    city_state_zip = block_lines[zip_idx]
    zip_code = ZIP_RE.search(city_state_zip).group(1)

    tokens = city_state_zip.split()
    state = ""
    if zip_code in tokens:
        z = tokens.index(zip_code)
        if z >= 1 and re.fullmatch(r"[A-Z]{2}", tokens[z - 1]):
            state = tokens[z - 1]

    city = ""
    if state and state in tokens:
        sidx = tokens.index(state)
        city = " ".join(tokens[:sidx]).strip()

    street = block_lines[zip_idx - 1] if zip_idx - 1 >= 0 else ""
    company = block_lines[0] if len(block_lines) >= 1 else ""

    # Clean company: avoid grabbing street
    if company == street and len(block_lines) >= 2:
        company = block_lines[1]

    out = {
        "Company": company.strip(),
        "Address": street.strip(),
        "City": city.strip(),
        "State": state.strip(),
        "ZipCode": zip_code.strip(),
        "Country": "USA",
    }
    return out


def extract_customer_number(words: List[dict], full_text: str) -> str:
    # coordinate-based first
    v = extract_value_near_label(words, ["Cust", "Cust#", "Customer", "Customer#", "CustomerNo", "Customer No"], page_num=1)
    v_digits = re.sub(r"[^\d]", "", v or "")
    if re.fullmatch(r"\d{3,10}", v_digits or ""):
        return v_digits

    # text fallback
    m = re.search(r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})", full_text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_created_by(words: List[dict], full_text: str) -> str:
    # coordinate-based label -> value
    v = extract_value_near_label(
        words,
        ["Created", "CreatedBy", "Created By", "Prepared", "PreparedBy", "Prepared By", "Entered", "EnteredBy", "Entered By"],
        page_num=1,
        prefer_right=True,
        right_max_tokens=10,
        below_lines=4,
    )

    # if value looks like it still contains "By", strip it
    v = re.sub(r"^\s*by\s+", "", (v or ""), flags=re.IGNORECASE).strip()

    # keep only name-ish tokens
    toks = v.split()
    name_parts = []
    for t in toks:
        if is_nameish(t):
            name_parts.append(t)
        else:
            break
    name = " ".join(name_parts).strip()
    if len(name.split()) >= 2:
        return name

    # full-text fallback (strong)
    m = re.search(r"(?:Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Z][A-Za-z.'-]+\s+[A-Z][A-Za-z.'-]+)", full_text, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # last fallback: known example
    m = re.search(r"\b(Loghan\s+Keefer)\b", full_text, re.IGNORECASE)
    if m:
        return "Loghan Keefer"

    return ""


def extract_salesperson(words: List[dict], full_text: str) -> str:
    # coordinate-based label -> value
    v = extract_value_near_label(
        words,
        ["Salesperson", "Sales", "SalesRep", "Sales Rep", "Sales Person", "Salesperson:"],
        page_num=1,
        prefer_right=True,
        right_max_tokens=10,
        below_lines=3,
    )
    v = clean_text(v)

    # pick name tokens
    toks = v.split()
    name_parts = []
    for t in toks:
        if is_nameish(t):
            name_parts.append(t)
        else:
            break
    name = " ".join(name_parts).strip()
    if name:
        return name

    # full-text fallback
    m = re.search(r"Salesperson\s*[:#]?\s*([A-Za-z.'-]+\s+[A-Za-z.'-]+)", full_text, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def extract_quote_number(full_text: str) -> str:
    m = QUOTE_NO_RE.search(full_text)
    return m.group(1).strip() if m else ""


def extract_quote_date(words: List[dict], full_text: str) -> str:
    # label based coordinate
    v = extract_value_near_label(words, ["Date", "QuoteDate", "Quote Date"], page_num=1, prefer_right=True, right_max_tokens=6, below_lines=3)
    d = DATE_RE.search(v or "")
    if d:
        return normalize_date(d.group(1))

    # fallback: first date in text
    m = DATE_RE.search(full_text)
    return normalize_date(m.group(1)) if m else ""


def extract_total_sales(full_text: str) -> str:
    m = re.search(r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})", full_text, re.IGNORECASE)
    if m:
        return normalize_money(m.group(1))
    monies = MONEY_RE.findall(full_text)
    return normalize_money(monies[-1]) if monies else ""


def extract_voelkr_fields(pdf_bytes: bytes) -> Dict[str, str]:
    words = extract_words_all_pages(pdf_bytes)
    full_text = extract_full_text(pdf_bytes)

    out: Dict[str, str] = {}
    out["Brand"] = VOELKR_DEFAULT_BRAND

    out["QuoteNumber"] = extract_quote_number(full_text)
    out["QuoteDate"] = extract_quote_date(words, full_text)
    out["CustomerNumber"] = extract_customer_number(words, full_text)

    # Salesperson -> ReferralManager
    out["ReferralManager"] = extract_salesperson(words, full_text)

    out["TotalSales"] = extract_total_sales(full_text)
    out["Created_By"] = extract_created_by(words, full_text)

    # Ship To -> Company + Address + City/State/Zip
    out.update(extract_ship_to_block(words))

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
