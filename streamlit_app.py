import io
import re
import zipfile
from datetime import datetime
from typing import List, Dict, Optional, Any, Tuple
from collections import defaultdict
from pathlib import Path

import pdfplumber
import pandas as pd
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
SYSTEM_TYPES = ["Cadre", "Voelkr"]

REPO_MAPPING_DIR = Path("mappings")
VOELKR_COLUMNS_PATH = REPO_MAPPING_DIR / "voelkr_columns.csv"
VOELKR_REGEX_MAP_PATH = REPO_MAPPING_DIR / "voelkr_regex_map.csv"


# =========================================================
# SECURITY / PRIVACY
# =========================================================
# - No PDF uploads are written to disk
# - No caching of uploaded bytes
# - Outputs are generated in-memory only
# =========================================================


# =========================================================
# COMMON HELPERS
# =========================================================
def normalize_date_str(date_str: Optional[str]) -> Optional[str]:
    if not date_str:
        return None
    m = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})", str(date_str))
    if not m:
        return str(date_str).strip()
    mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:
        yy += 2000
    return f"{mm:02d}/{dd:02d}/{yy:04d}"


def extract_full_text(pdf_bytes: bytes) -> str:
    full_text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                txt = page.extract_text(layout=True) or ""
            except TypeError:
                txt = page.extract_text() or ""
            full_text += txt + "\n"
    return full_text


def safe_float(x: Optional[str]) -> Optional[float]:
    if x is None:
        return None
    s = str(x).replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None


# =========================================================
# LOAD VOELKR HEADERS (from repo)
# =========================================================
def load_voelkr_columns() -> List[str]:
    if not VOELKR_COLUMNS_PATH.exists():
        raise FileNotFoundError(
            f"Missing {VOELKR_COLUMNS_PATH}. Add your uploaded CSV header row to that file in your GitHub repo."
        )
    # read only header row
    df0 = pd.read_csv(VOELKR_COLUMNS_PATH, nrows=0)
    cols = list(df0.columns)
    if not cols:
        raise ValueError(f"{VOELKR_COLUMNS_PATH} has no columns/header row.")
    return cols


# =========================================================
# LOAD VOELKR REGEX MAP (from repo)
# =========================================================
def load_voelkr_regex_map() -> pd.DataFrame:
    if not VOELKR_REGEX_MAP_PATH.exists():
        # Not fatal; app can still run and output blanks
        return pd.DataFrame(columns=["field", "regex", "type", "default"])
    df = pd.read_csv(VOELKR_REGEX_MAP_PATH)
    # normalize columns
    df.columns = [c.strip().lower() for c in df.columns]
    for need in ["field", "regex"]:
        if need not in df.columns:
            raise ValueError(f"{VOELKR_REGEX_MAP_PATH} must contain columns: field, regex (and optionally type, default)")
    if "type" not in df.columns:
        df["type"] = ""
    if "default" not in df.columns:
        df["default"] = ""
    return df


def apply_cast(value: str, typ: str) -> str:
    v = (value or "").strip()
    typ = (typ or "").strip().lower()
    if not v:
        return ""
    if typ in ("date", "mm/dd/yyyy"):
        return normalize_date_str(v) or v
    if typ in ("number", "float", "amount", "currency"):
        # return numeric string (not float) to avoid rounding issues
        s = v.replace("$", "").replace(",", "")
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return m.group(0) if m else v
    return v


# =========================================================
# VOELKR EXTRACTOR (HIGH ACCURACY VIA REGEX MAP)
# =========================================================
def build_rows_voelkr(pdf_bytes: bytes, filename: str, voelkr_cols: List[str], map_df: pd.DataFrame) -> List[Dict[str, Any]]:
    text = extract_full_text(pdf_bytes)

    row: Dict[str, Any] = {c: "" for c in voelkr_cols}

    # Ensure we always keep filename somewhere
    if "PDF" in row:
        row["PDF"] = filename
    else:
        row["PDF"] = filename  # even if not in columns, we’ll add later

    # Apply regex map
    for _, r in map_df.iterrows():
        field = str(r.get("field", "")).strip()
        regex = str(r.get("regex", "")).strip()
        typ = str(r.get("type", "")).strip()
        default = str(r.get("default", "")).strip()

        if not field or not regex:
            continue

        value = ""
        try:
            m = re.search(regex, text, flags=re.IGNORECASE | re.MULTILINE)
            if m:
                value = m.group(1) if m.lastindex else m.group(0)
                value = re.sub(r"\s+", " ", str(value)).strip()
        except re.error:
            value = ""

        if not value:
            value = default

        value = apply_cast(value, typ)

        # Only fill if field exists in your predefined headers
        if field in row:
            row[field] = value

    return [row]


# =========================================================
# CADRE (keep your existing working Cadre logic here)
# NOTE: I’m leaving Cadre as “stub” so this file runs.
# Replace build_rows_cadre() with your existing Cadre parser.
# =========================================================
def build_rows_cadre(pdf_bytes: bytes, filename: str) -> List[Dict[str, Any]]:
    # TODO: paste your existing Cadre extraction here
    return [{
        "PDF": filename,
        "item_id": "",
        "item_desc": "",
    }]


# =========================================================
# ROUTER
# =========================================================
def parse_pdf(system_type: str, pdf_bytes: bytes, filename: str,
              voelkr_cols: Optional[List[str]] = None,
              voelkr_map: Optional[pd.DataFrame] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
    warnings: List[str] = []
    if system_type == "Cadre":
        return build_rows_cadre(pdf_bytes, filename), warnings

    if system_type == "Voelkr":
        if voelkr_cols is None or voelkr_map is None:
            raise ValueError("Voelkr mapping not loaded.")
        return build_rows_voelkr(pdf_bytes, filename, voelkr_cols, voelkr_map), warnings

    raise ValueError(f"Unknown system type: {system_type}")


# =========================================================
# STREAMLIT UI
# =========================================================
st.set_page_config(page_title="PDF Extractor", layout="wide")
st.title("PDF Extractor")

st.markdown(
    """
- Select **System type** (Cadre / Voelkr)
- Upload PDF(s)
- Click **Extract**
- Download **one ZIP** (CSV + PDFs)

**Privacy:** Files are processed in-memory and not saved by this app.
"""
)

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "result_zip_bytes" not in st.session_state:
    st.session_state["result_zip_bytes"] = None
if "result_preview" not in st.session_state:
    st.session_state["result_preview"] = None
if "result_summary" not in st.session_state:
    st.session_state["result_summary"] = None

system_type = st.selectbox("System type", SYSTEM_TYPES, index=0)
debug_mode = st.checkbox("Debug mode", value=False)

uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"pdf_uploader_{st.session_state['uploader_key']}",
)

extract_btn = st.button("Extract")

# Preload Voelkr mapping ONCE (no repeated uploads)
voelkr_cols: Optional[List[str]] = None
voelkr_map: Optional[pd.DataFrame] = None
if system_type == "Voelkr":
    try:
        voelkr_cols = load_voelkr_columns()
        voelkr_map = load_voelkr_regex_map()
        if debug_mode:
            st.caption(f"Loaded Voelkr columns: {len(voelkr_cols)}")
            st.caption(f"Loaded Voelkr regex rules: {len(voelkr_map)}")
    except Exception as e:
        st.error(str(e))

# show results if present
if st.session_state["result_zip_bytes"] is not None:
    st.success(st.session_state["result_summary"] or "Done.")
    if st.session_state["result_preview"] is not None:
        st.subheader("Preview (first 50 rows)")
        st.dataframe(st.session_state["result_preview"], use_container_width=True)

    st.download_button(
        "Download ZIP (CSV + PDFs)",
        data=st.session_state["result_zip_bytes"],
        file_name=f"{system_type.lower()}_extraction_output.zip",
        mime="application/zip",
    )

    if st.button("New extraction"):
        st.session_state["result_zip_bytes"] = None
        st.session_state["result_preview"] = None
        st.session_state["result_summary"] = None
        st.session_state["uploader_key"] += 1
        st.rerun()

if extract_btn:
    if not uploaded_files:
        st.error("Please upload at least one PDF.")
    elif system_type == "Voelkr" and (voelkr_cols is None or voelkr_map is None):
        st.error("Voelkr mapping is not loaded. Add mappings/voelkr_columns.csv and (optionally) mappings/voelkr_regex_map.csv to your repo.")
    else:
        file_data = [{"name": f.name, "bytes": f.read()} for f in uploaded_files]

        all_rows: List[Dict[str, Any]] = []
        warnings_all: List[str] = []

        progress = st.progress(0.0)
        status = st.empty()

        for idx, fd in enumerate(file_data, start=1):
            status.text(f"Processing {idx}/{len(file_data)}: {fd['name']}")
            rows, warns = parse_pdf(system_type, fd["bytes"], fd["name"], voelkr_cols, voelkr_map)
            all_rows.extend(rows)
            warnings_all.extend([f"{fd['name']}: {w}" for w in warns])
            progress.progress(idx / len(file_data))

        df = pd.DataFrame(all_rows)

        # Ensure final columns match “predefined headers”
        if system_type == "Voelkr" and voelkr_cols:
            for c in voelkr_cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[voelkr_cols]
        else:
            # Cadre: keep your own ordering when you paste it in
            pass

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("extracted/extracted.csv", df.to_csv(index=False).encode("utf-8"))
            for fd in file_data:
                zf.writestr(f"pdfs/{fd['name']}", fd["bytes"])
            if warnings_all:
                zf.writestr("extracted/warnings.txt", "\n".join(warnings_all).encode("utf-8"))
        zip_buf.seek(0)

        st.session_state["result_zip_bytes"] = zip_buf.getvalue()
        st.session_state["result_preview"] = df.head(50)
        st.session_state["result_summary"] = f"Parsed {len(file_data)} PDF(s). Output rows: {len(df)}."

        # Clear uploader after extraction
        st.session_state["uploader_key"] += 1
        st.rerun()
