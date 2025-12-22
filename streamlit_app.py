import io
import re
import zipfile
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import pdfplumber
import streamlit as st


# =========================
# SECURITY / PRIVACY NOTES
# =========================
# - This app does NOT write uploads to disk.
# - Processing is done in memory only.
# - Avoid st.cache_data / st.cache_resource for uploaded content.
# - If deployed on Streamlit Community Cloud, files exist only for the session runtime.


# =========================
# UTILITIES
# =========================
def _normalize_label(s: str) -> str:
    """Normalize label text for matching."""
    s = (s or "").strip().lower()
    s = re.sub(r"[\s\t]+", " ", s)
    s = re.sub(r"[^a-z0-9 :/_\-().]", "", s)
    return s


def _clean_value(s: str) -> str:
    """Clean extracted value."""
    s = (s or "").strip()
    # collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract text from all pages with best-effort layout preservation."""
    text_parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # layout=True helps in many PDFs; fallback if not supported by version
            try:
                t = page.extract_text(layout=True) or ""
            except TypeError:
                t = page.extract_text() or ""
            text_parts.append(t)
    return "\n".join(text_parts)


def extract_pdf_words(pdf_bytes: bytes) -> List[dict]:
    """Extract words (with coordinates) from all pages."""
    words_all: List[dict] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=True,
            ) or []
            # keep page number for debugging / advanced matching
            for w in words:
                w["page_number"] = page.page_number
            words_all.extend(words)
    return words_all


def build_label_value_index_from_text(full_text: str) -> Dict[str, str]:
    """
    Build a normalized label->value dict from patterns like:
      Label: Value
      Label ..... Value
    This supports lots of real-world PDFs.
    """
    idx: Dict[str, str] = {}
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    # Common patterns
    # 1) "Label: Value"
    colon_re = re.compile(r"^(.{2,60}?):\s*(.+)$")
    # 2) "Label    Value" where spacing is large (dot leaders or multiple spaces)
    space_re = re.compile(r"^(.{2,60}?)(?:\s{3,}|\s\.\.\.\s+|\s\.\s+)(.+)$")

    for ln in lines:
        m = colon_re.match(ln)
        if m:
            label = _normalize_label(m.group(1))
            val = _clean_value(m.group(2))
            if label and val and label not in idx:
                idx[label] = val
            continue

        m = space_re.match(ln)
        if m:
            label = _normalize_label(m.group(1))
            val = _clean_value(m.group(2))
            if label and val and label not in idx:
                idx[label] = val

    return idx


# =========================
# MAPPING CSV LOGIC
# =========================
def read_mapping_csv(mapping_csv_bytes: bytes) -> Tuple[List[str], Optional[pd.DataFrame]]:
    """
    Supports two mapping CSV types:

    A) "Header-only CSV": one row or just columns representing the desired output fields.
       Example: columns = ["QuoteNumber","QuoteDate","Company",...]

    B) "Explicit mapping CSV" (recommended for accuracy):
       Must include a column called 'field' plus one of:
         - 'regex' (preferred), OR
         - 'label' (exact label text in PDF)

       Optional columns:
         - 'type' (date, number, text)
         - 'postprocess_regex' (to extract subgroup)
         - 'default'
    """
    df = pd.read_csv(io.BytesIO(mapping_csv_bytes))

    cols_norm = [_normalize_label(c) for c in df.columns]

    if "field" in cols_norm and ("regex" in cols_norm or "label" in cols_norm):
        # explicit mapping table
        # normalize column names access
        df.columns = cols_norm
        # output fields in explicit order
        fields = df["field"].astype(str).tolist()
        return fields, df

    # header-only: output fields are the CSV columns
    fields = list(df.columns)
    return fields, None


def apply_type_cast(value: str, typ: str) -> str:
    v = (value or "").strip()
    typ = (typ or "").strip().lower()
    if not v:
        return ""

    if typ in ("date", "mm/dd/yyyy", "mm-dd-yyyy"):
        # Try to standardize mm/dd/yyyy if it looks like a date
        m = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})", v)
        if m:
            mm = int(m.group(1))
            dd = int(m.group(2))
            yy = int(m.group(3))
            if yy < 100:
                yy += 2000
            return f"{mm:02d}/{dd:02d}/{yy:04d}"
        return v

    if typ in ("number", "float", "amount", "currency"):
        # keep as clean numeric string (don’t force float if you want exact)
        v2 = v.replace(",", "")
        v2 = v2.replace("$", "")
        # keep only first numeric pattern
        m = re.search(r"-?\d+(?:\.\d+)?", v2)
        return m.group(0) if m else v

    return v


def extract_fields_from_pdf(
    pdf_bytes: bytes,
    output_fields: List[str],
    mapping_df: Optional[pd.DataFrame],
) -> Dict[str, str]:
    """
    Main extraction:
    - Always extracts full text
    - Builds label:value index for fast lookup
    - If mapping_df provided:
        - use regex per field OR label per field
      else:
        - auto-match using field name as label
    """
    full_text = extract_pdf_text(pdf_bytes)
    label_index = build_label_value_index_from_text(full_text)

    results: Dict[str, str] = {}

    if mapping_df is None:
        # Auto: use output field name as label key
        for f in output_fields:
            key = _normalize_label(str(f))
            results[f] = label_index.get(key, "")
        return results

    # Explicit mapping
    for _, row in mapping_df.iterrows():
        field = str(row.get("field", "")).strip()
        if not field:
            continue

        default_val = str(row.get("default", "")).strip() if "default" in mapping_df.columns else ""
        typ = str(row.get("type", "")).strip() if "type" in mapping_df.columns else ""

        value = ""

        # Preferred: regex
        regex = str(row.get("regex", "")).strip() if "regex" in mapping_df.columns else ""
        if regex:
            try:
                m = re.search(regex, full_text, flags=re.IGNORECASE | re.MULTILINE)
                if m:
                    # If regex has capturing group(s), use group(1)
                    value = m.group(1) if m.lastindex else m.group(0)
                    value = _clean_value(value)
            except re.error:
                value = ""

        # Fallback: label lookup
        if not value:
            label = str(row.get("label", "")).strip() if "label" in mapping_df.columns else ""
            if label:
                value = label_index.get(_normalize_label(label), "")

        # Optional postprocess regex (extract subpart)
        if value and "postprocess_regex" in mapping_df.columns:
            post = str(row.get("postprocess_regex", "")).strip()
            if post:
                try:
                    m2 = re.search(post, value, flags=re.IGNORECASE)
                    if m2:
                        value = m2.group(1) if m2.lastindex else m2.group(0)
                        value = _clean_value(value)
                except re.error:
                    pass

        if not value:
            value = default_val

        value = apply_type_cast(value, typ)
        results[field] = value

    # Ensure all output fields exist
    for f in output_fields:
        if f not in results:
            results[f] = ""

    return results


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="PDF Header Field Extractor", layout="wide")
st.title("PDF Header Field Extractor (Mapping CSV → Output)")

st.markdown(
    """
Upload:
1) **Mapping CSV** (your field list / mapping table)  
2) One or more **PDFs**

Then download **one ZIP** containing:
- extracted CSV
- uploaded PDFs
"""
)

# session state for clearing uploader + keeping output
if "uploader_key2" not in st.session_state:
    st.session_state["uploader_key2"] = 0
if "result_zip2" not in st.session_state:
    st.session_state["result_zip2"] = None
if "result_preview2" not in st.session_state:
    st.session_state["result_preview2"] = None
if "result_summary2" not in st.session_state:
    st.session_state["result_summary2"] = None

system_type = st.selectbox("System type", ["Cadre (Header Mapping Tool)"], index=0)
debug_mode = st.checkbox("Debug mode", value=False)

mapping_file = st.file_uploader(
    "Upload mapping CSV (required)",
    type=["csv"],
    accept_multiple_files=False,
    key=f"mapping_uploader_{st.session_state['uploader_key2']}",
)

uploaded_pdfs = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader2_{st.session_state['uploader_key2']}",
)

extract_btn = st.button("Extract")

# Show results if available
if st.session_state["result_zip2"] is not None:
    st.success(st.session_state["result_summary2"] or "Done.")
    if st.session_state["result_preview2"] is not None:
        st.subheader("Preview (first 50 rows)")
        st.dataframe(st.session_state["result_preview2"], use_container_width=True)

    st.download_button(
        "Download ZIP (CSV + PDFs)",
        data=st.session_state["result_zip2"],
        file_name="header_extraction_output.zip",
        mime="application/zip",
    )

    if st.button("New extraction"):
        st.session_state["result_zip2"] = None
        st.session_state["result_preview2"] = None
        st.session_state["result_summary2"] = None
        st.session_state["uploader_key2"] += 1
        st.rerun()

if extract_btn:
    if mapping_file is None:
        st.error("Please upload the mapping CSV.")
    elif not uploaded_pdfs:
        st.error("Please upload at least one PDF.")
    else:
        mapping_bytes = mapping_file.read()
        output_fields, mapping_df = read_mapping_csv(mapping_bytes)

        # Read all PDFs
        file_data = [{"name": f.name, "bytes": f.read()} for f in uploaded_pdfs]

        all_rows: List[Dict[str, str]] = []
        progress = st.progress(0.0)
        status = st.empty()

        for idx, fd in enumerate(file_data, start=1):
            status.text(f"Processing {idx}/{len(file_data)}: {fd['name']}")
            row = extract_fields_from_pdf(fd["bytes"], output_fields, mapping_df)
            row["PDF"] = fd["name"]  # keep traceability
            all_rows.append(row)
            progress.progress(idx / len(file_data))

            if debug_mode:
                st.write(f"DEBUG: {fd['name']} extracted keys:", {k: row.get(k, "") for k in output_fields[:10]})

        df = pd.DataFrame(all_rows)
        # Ensure column order: output_fields + PDF at end (if PDF isn't already in mapping)
        cols = [c for c in output_fields if c in df.columns]
        if "PDF" in df.columns and "PDF" not in cols:
            cols.append("PDF")
        df = df.reindex(columns=cols)

        # Build ZIP: extracted CSV + PDFs + mapping copy
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("extracted/extracted_header_fields.csv", df.to_csv(index=False).encode("utf-8"))
            zf.writestr("mapping/mapping.csv", mapping_bytes)
            for fd in file_data:
                zf.writestr(f"pdfs/{fd['name']}", fd["bytes"])
        zip_buf.seek(0)

        st.session_state["result_zip2"] = zip_buf.getvalue()
        st.session_state["result_preview2"] = df.head(50)
        st.session_state["result_summary2"] = f"Parsed {len(file_data)} PDF(s). Output rows: {len(df)}."

        # Clear uploaders
        st.session_state["uploader_key2"] += 1
        st.rerun()
