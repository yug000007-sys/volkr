import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# -----------------------------
# PAGE SETUP (prevents blank page)
# -----------------------------
st.set_page_config(page_title="Voelker PDF to Excel", layout="wide")
st.title("Voelker PDF → Excel Automation")
st.caption("Upload Voelker quote PDFs and download structured Excel output")

# -----------------------------
# HELPERS
# -----------------------------
def extract_text(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
    return text

def find(pattern, text):
    m = re.search(pattern, text, re.MULTILINE)
    return m.group(1).strip() if m else ""

def parse_pdf(text):
    data = {}

    data["ReferralManager"] = find(r"\d{2}/\d{2}/\d{2}\s+\d{2}/\d{2}/\d{2}.*\s([A-Z]+)\s+\d+", text)
    data["CustomerNumber"] = find(r"\s(\d{4,})\s+NET", text)
    data["QuoteDate"] = find(r"(\d{1,2}/\d{1,2}/\d{2})\s+\d{1,2}/\d{1,2}/\d{2}", text)
    data["QuoteNumber"] = find(r"(0\d/\d{6})", text)
    data["Created_By"] = find(r"(0\d/\d{6})\s+([A-Za-z ]+)\s+UPS", text)
    data["TotalSales"] = find(r"Quote Total\s+([\d,]+\.\d{2})", text).replace(",", "")

    # Ship To parsing
    ship = re.search(
        r"Ship To\s*\n([A-Z0-9 &]+)\n([A-Z0-9 .]+)\n([A-Z ]+)\s+([A-Z]{2})\s+(\d{5})",
        text,
        re.MULTILINE
    )

    if ship:
        data["Company"] = ship.group(1)
        data["Address"] = ship.group(2)
        data["City"] = ship.group(3)
        data["State"] = ship.group(4)
        data["Zip"] = ship.group(5)
    else:
        data["Company"] = ""
        data["Address"] = ""
        data["City"] = ""
        data["State"] = ""
        data["Zip"] = ""

    return data

# -----------------------------
# UI
# -----------------------------
uploaded_files = st.file_uploader(
    "Upload Voelker PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} PDF(s) uploaded")

    if st.button("Generate Excel"):
        rows = []

        for pdf in uploaded_files:
            text = extract_text(pdf)
            parsed = parse_pdf(text)

            row = {
                "PDFName": pdf.name,
                "ReferralManager": parsed.get("ReferralManager", ""),
                "QuoteNumber": parsed.get("QuoteNumber", ""),
                "QuoteDate": parsed.get("QuoteDate", ""),
                "Company": parsed.get("Company", ""),
                "Address": parsed.get("Address", ""),
                "City": parsed.get("City", ""),
                "State": parsed.get("State", ""),
                "Zip": parsed.get("Zip", ""),
                "TotalSales": parsed.get("TotalSales", ""),
                "Created_By": parsed.get("Created_By", ""),
                "CustomerNumber": parsed.get("CustomerNumber", "")
            }

            rows.append(row)

        df = pd.DataFrame(rows)

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        # Excel download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="VoelkerQuotes")

        st.download_button(
            label="Download Excel",
            data=buffer.getvalue(),
            file_name="voelker_quotes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Upload PDF files to begin.")
