import io
import re
import zipfile
from typing import List, Dict, Any
from collections import defaultdict

import pdfplumber
import pandas as pd
import streamlit as st


# ===============================
# CONFIG
# ===============================
SYSTEM_TYPES = ["Cadre", "Voelkr"]
VOELKR_BRAND = "Voelker Controls"

COLUMNS = [
    "ReferralManager","ReferralEmail","QuoteNumber","QuoteDate","Company",
    "FirstName","LastName","ContactEmail","ContactPhone","Address","County",
    "City","State","ZipCode","Country","manufacturer_Name","item_id","item_desc",
    "Quantity","TotalSales","PDF","Brand","QuoteExpiration","CustomerNumber",
    "UnitSales","Unit_Cost","sales_cost","cust_type","QuoteComment",
    "Created_By","quote_line_no","DemoQuote"
]


# ===============================
# PDF HELPERS
# ===============================
def extract_words(pdf_bytes):
    words = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            for w in p.extract_words(x_tolerance=2, y_tolerance=2, use_text_flow=True):
                w["page"] = p.page_number
                words.append(w)
    return words


def group_rows(words, page=1, tol=3):
    rows = defaultdict(list)
    for w in words:
        if w["page"] == page:
            key = round(w["top"] / tol) * tol
            rows[key].append(w)
    return [sorted(rows[k], key=lambda x: x["x0"]) for k in sorted(rows)]


def row_text(row):
    return " ".join(w["text"] for w in row).strip()


# ===============================
# FIELD EXTRACTORS
# ===============================
def extract_quote_number(text):
    m = re.search(r"\b\d{2}/\d{6}\b", text)
    return m.group(0) if m else ""


def extract_quote_date(text):
    m = re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", text)
    return m.group(0) if m else ""


def extract_total(text):
    m = re.search(r"\b([\d,]+\.\d{2})\b", text)
    return m.group(1) if m else ""


def extract_customer_number(words):
    for row in group_rows(words):
        for i, w in enumerate(row):
            if w["text"].lower().startswith("cust"):
                for x in row[i+1:i+8]:
                    v = re.sub(r"\D", "", x["text"])
                    if v.isdigit():
                        return v
    return ""


def extract_created_by(words, text):
    m = re.search(r"(Created By|Prepared By)\s+([A-Z][a-z]+\s+[A-Z][a-z]+)", text)
    if m:
        return m.group(2)
    for row in group_rows(words):
        t = row_text(row).lower()
        if "created" in t or "prepared" in t:
            for w in row:
                if re.fullmatch(r"[A-Z][a-z]+", w["text"]):
                    return w["text"]
    return ""


def extract_salesperson(words):
    for row in group_rows(words):
        t = row_text(row).lower()
        if "sales" in t:
            for w in row:
                if re.fullmatch(r"[A-Z][A-Z]+", w["text"]):
                    return w["text"]
    return ""


def extract_ship_to(words):
    rows = group_rows(words)
    for i, r in enumerate(rows):
        if "ship to" in row_text(r).lower():
            block = []
            for j in range(i+1, i+8):
                if j < len(rows):
                    block.append(row_text(rows[j]))
            break
    else:
        return {}

    # STRICT POSITIONS
    company = block[0]
    address = block[1]
    city_state_zip = block[2]

    m = re.search(r"(.*)\s+([A-Z]{2})\s+(\d{5})", city_state_zip)
    city, state, zipc = ("","","")
    if m:
        city, state, zipc = m.group(1), m.group(2), m.group(3)

    return {
        "Company": company,
        "Address": address,
        "City": city,
        "State": state,
        "ZipCode": zipc,
        "Country": "USA",
    }


# ===============================
# MAIN PARSER
# ===============================
def parse_voelkr(pdf_bytes, filename):
    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            text += p.extract_text() or ""

    words = extract_words(pdf_bytes)

    out = {c: "" for c in COLUMNS}
    out["Brand"] = VOELKR_BRAND
    out["PDF"] = filename

    out["QuoteNumber"] = extract_quote_number(text)
    out["QuoteDate"] = extract_quote_date(text)
    out["TotalSales"] = extract_total(text)
    out["CustomerNumber"] = extract_customer_number(words)
    out["Created_By"] = extract_created_by(words, text)
    out["ReferralManager"] = extract_salesperson(words)

    out.update(extract_ship_to(words))

    return out


# ===============================
# STREAMLIT UI
# ===============================
st.set_page_config(layout="wide")
st.title("Voelkr PDF Extractor")

system = st.selectbox("System", SYSTEM_TYPES, index=1)

files = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)

if st.button("Extract"):
    rows = []
    for f in files:
        rows.append(parse_voelkr(f.read(), f.name))

    df = pd.DataFrame(rows, columns=COLUMNS)
    st.dataframe(df)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("voelkr.csv", df.to_csv(index=False))
        for f in files:
            z.writestr(f.name, f.getvalue())
    buf.seek(0)

    st.download_button("Download ZIP", buf, "voelkr_output.zip")
