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
ZIP5_RE = re.compile(r"\b(\d{5})(?:-\d{4})?\b")
CITY_ST_ZIP_RE = re.compile(r"^(.*?)\s+([A-Z]{2})\s+(\d{5})(?:-\d{4})?$")
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
                # normalized text for token matching
                w["text_norm"] = (
                    (w.get("text") or "")
                    .strip()
                    .lower()
                    .replace(":", "")
                    .replace("#", "")
                )
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
# VOELKR EXTRACTORS
# =========================================================
def extract_quote_number(full_text: str) -> str:
    m = QUOTE_NO_RE.search(full_text)
    return m.group(1) if m else ""


def extract_quote_date(full_text: str) -> str:
    m = DATE_RE.search(full_text)
    return normalize_date(m.group(0)) if m else ""


def extract_total_sales(full_text: str) -> str:
    m = re.search(
        r"(?:Total\s*Sales|Grand\s*Total|Total)\s*[: ]+\$?\s*([\d,]+\.\d{2})",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return m.group(1)
    monies = MONEY_RE.findall(full_text)
    return monies[-1] if monies else ""


def find_ship_to_anchor(rows: List[Dict[str, Any]]) -> Optional[dict]:
    """
    Match 'Ship To' even if split into separate tokens or has punctuation.
    """
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


def _looks_like_street(s: str) -> bool:
    s = clean_text(s)
    return bool(re.match(r"^\d{2,6}\s+", s))


def _looks_like_country_only(s: str) -> bool:
    s = clean_text(s).upper()
    return s in {"USA", "UNITED STATES", "UNITED STATES OF AMERICA"}


def _parse_city_state_zip(line: str) -> Dict[str, str]:
    """
    Handles:
      'TRENTON OH 45067-9760'
      'TIPP CITY OH 45371'
    """
    line = clean_text(line)
    # normalize "City, ST ZIP" -> "City ST ZIP"
    line = line.replace(",", " ")
    line = re.sub(r"\s{2,}", " ", line).strip()

    # Find ZIP (5 or 9)
    mzip = re.search(r"\b(\d{5})(?:-(\d{4}))?\b", line)
    if not mzip:
        return {"City": "", "State": "", "ZipCode": ""}

    zip5 = mzip.group(1)
    # take the token just before zip as state, and everything before that as city
    before = line[: mzip.start()].strip()
    parts = before.split()
    if len(parts) < 2:
        return {"City": before, "State": "", "ZipCode": zip5}

    state = parts[-1]
    city = " ".join(parts[:-1]).strip()

    # guard: state should be 2 letters
    if not re.fullmatch(r"[A-Z]{2}", state.upper()):
        # fallback to regex City ST ZIP
        m = CITY_ST_ZIP_RE.match(re.sub(r"\b(\d{5})(?:-\d{4})?\b", zip5, line))
        if m:
            return {"City": clean_text(m.group(1)), "State": m.group(2), "ZipCode": zip5}
        return {"City": before, "State": "", "ZipCode": zip5}

    return {"City": clean_text(city), "State": state.upper(), "ZipCode": zip5}


def extract_ship_to_block(words: List[dict]) -> Dict[str, str]:
    """
    Required behavior:
      Company/Address/City/State/Zip MUST come from Ship To.

    Robust rules based on your sample:
      Ship To:
        TRENTON BREWERY
        2525 WAYNE MADISON ROAD
        TRENTON OH 45067-9760

    Implementation:
      1) find Ship To anchor row
      2) take next non-empty lines in the Ship To column window
      3) Company = first meaningful line (not street, not city/state/zip, not country)
      4) Address = first street-like line after company
      5) City/State/Zip = first line containing ZIP after address/company
    """
    rows = group_rows(words, page=1, y_tol=3.0)
    anchor = find_ship_to_anchor(rows)
    if not anchor:
        return {}

    start_i = anchor["row_idx"] + 1
    col_left = max(0, anchor["x0"] - 160)
    col_right = anchor["x1"] + 900

    # collect candidate lines under Ship To
    lines: List[str] = []
    for j in range(start_i, min(start_i + 30, len(rows))):
        row_words = rows[j]["words"]
        in_col = [w for w in row_words if w["x0"] >= col_left and w["x1"] <= col_right]
        line = clean_text(" ".join(w["text"] for w in in_col))
        if not line:
            continue

        # stop when leaving ship-to block
        if re.search(r"\b(bill\s*to|sold\s*to|subtotal|grand\s*total|total\b|items?|product)\b", line, re.IGNORECASE):
            break

        lines.append(line)

    if not lines:
        return {}

    # identify city/state/zip line index = first line with ZIP
    zip_i = None
    for i, ln in enumerate(lines):
        if ZIP5_RE.search(ln):
            zip_i = i
            break

    # Company = first meaningful line before zip_i
    company = ""
    company_i = None
    search_upto = zip_i if zip_i is not None else len(lines)
    for i in range(0, search_upto):
        ln = lines[i].strip()
        if not ln:
            continue
        if _looks_like_country_only(ln):
            continue
        if _looks_like_street(ln):
            continue
        if ZIP5_RE.search(ln):
            continue
        # avoid accidentally taking "Ship To" itself if it leaks
        if ln.lower().startswith("ship"):
            continue
        company = ln
        company_i = i
        break

    # Address = first street line after company (or from top if company missing)
    address = ""
    start_addr = (company_i + 1) if company_i is not None else 0
    end_addr = zip_i if zip_i is not None else len(lines)
    for i in range(start_addr, end_addr):
        ln = lines[i].strip()
        if _looks_like_street(ln):
            address = ln
            break

    # City/State/Zip from zip_i line
    city = state = zipc = ""
    if zip_i is not None:
        parsed = _parse_city_state_zip(lines[zip_i])
        city, state, zipc = parsed["City"], parsed["State"], parsed["ZipCode"]

    # clean duplicates in address (common extraction glitch)
    if address:
        toks = address.split()
        if len(toks) >= 4:
            half = len(toks) // 2
            if toks[:half] == toks[half:]:
                address = " ".join(toks[:half])

    return {
        "Company": company,
        "Address": address,
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
            cust_words = [
                w for w in r["words"]
                if w["text_norm"].startswith("cust") or w["text_norm"].startswith("customer")
            ]
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

    m = re.search(
        r"(?:cust|customer)\s*(?:#|no\.?|number)?\s*[:#\-]?\s*(\d{3,10})",
        full_text,
        re.IGNORECASE,
    )
    return m.group(1) if m else ""


def extract_created_by(words: List[dict], full_text: str) -> str:
    m = re.search(
        r"(Created\s*By|Prepared\s*By|Entered\s*By)\s*[:#]?\s*([A-Z][A-Za-z.'-]+\s+[A-Z][A-Za-z.'-]+)",
        full_text,
        re.IGNORECASE,
    )
    if m:
        return clean_text(m.group(2))
    return ""


def extract_referral_manager(words: List[dict], full_text: str) -> str:
    """
    IMPORTANT: No guessing.
    Only use explicit labels for salesperson/sales rep; otherwise blank.

    Acceptable labels (case-insensitive):
      Salesperson
      Sales Rep
      Sales Rep.
      Salesperson:
    """
    rows = group_rows(words, page=1, y_tol=3.0)

    label_re = re.compile(r"\b(sales\s*rep|salesperson)\b", re.IGNORECASE)

    for r in rows:
        if not label_re.search(r["text"]):
            continue

        # take text to the right of the label on the same row
        # heuristic: last ALLCAPS token (like DAYTON) OR last word chunk after label
        allcaps = []
        for w in r["words"]:
            txt = w["text"].strip()
            if re.fullmatch(r"[A-Z]{3,}", txt) and txt.lower() not in {"cust", "customer"}:
                allcaps.append(txt)
        if allcaps:
            return allcaps[-1]

        # fallback: take last two "name-like" words after label
        parts = [w["text"].strip() for w in r["words"] if w["text"].strip()]
        # remove obvious label words
        cleaned = [p for p in parts if p.lower() not in {"sales", "rep", "salesrep", "salesperson"}]
        # return last token if any
        return cleaned[-1] if cleaned else ""

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

    out.update(extract_ship_to_block(words))
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
