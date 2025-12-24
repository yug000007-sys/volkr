import re
import csv
from pathlib import Path

import pdfplumber


# ---------- helpers ----------
def clean(s: str) -> str:
    return re.sub(r"[ \t]+", " ", (s or "").strip())

def find_labeled_value(text: str, label: str) -> str:
    """
    Looks for patterns like:
      Label: value
      Label value
    Captures up to end of line.
    """
    # Try "Label: value"
    m = re.search(rf"(?im)^\s*{re.escape(label)}\s*:\s*(.+?)\s*$", text)
    if m:
        return clean(m.group(1))

    # Try "Label   value" (2+ spaces between)
    m = re.search(rf"(?im)^\s*{re.escape(label)}\s{{2,}}(.+?)\s*$", text)
    if m:
        return clean(m.group(1))

    return ""

def extract_ship_to(text: str):
    """
    CRITICAL: Always extract address fields from the “Ship To” section.
    Assumes Ship To block like:

      Ship To:
      COMPANY NAME
      123 MAIN ST
      SUITE 400   (optional)
      CITY, ST 12345

    Returns: (company, street_address, city, state, zip)
    """
    # Grab lines and locate the Ship To header line
    lines = [clean(l) for l in text.splitlines()]
    ship_idx = None
    for i, line in enumerate(lines):
        if re.fullmatch(r"(?i)ship\s*to\s*:?", line):
            ship_idx = i
            break

    if ship_idx is None:
        return ("", "", "", "", "")

    # Collect following non-empty lines until a stopping condition
    # Stop if we hit common next-section headers (adjustable)
    stop_headers = re.compile(
        r"(?i)^(bill\s*to|sold\s*to|remit\s*to|terms|notes|ship\s*via|quote|customer|cust\s*#|salesperson|quoted\s*by)\b"
    )

    block = []
    for j in range(ship_idx + 1, len(lines)):
        if not lines[j]:
            # allow a single blank inside, but break on multiple blanks after some content
            if block:
                # peek ahead: if next non-empty is a header, stop
                continue
            else:
                continue
        if stop_headers.match(lines[j]) and block:
            break
        # Also stop if we hit another label style line like "X: Y" after collecting something
        if block and re.match(r"(?i)^[A-Za-z][A-Za-z \/#&\.-]{2,}:\s*\S+", lines[j]):
            break
        block.append(lines[j])
        # safety: don't let it run too far
        if len(block) >= 8:
            break

    if not block:
        return ("", "", "", "", "")

    company = block[0] if len(block) >= 1 else ""

    # Find the city/state/zip line (usually last)
    city = state = zip_code = ""
    city_state_zip_idx = None

    # Try from bottom up to find "City, ST 12345" OR "City ST 12345"
    for k in range(len(block) - 1, -1, -1):
        m = re.match(r"^(?P<city>.+?)[,\s]+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)$", block[k])
        if m:
            city = clean(m.group("city").rstrip(","))
            state = m.group("state")
            zip_code = m.group("zip")
            city_state_zip_idx = k
            break

    # Street address = lines between company and city/state/zip line
    street_lines = []
    if city_state_zip_idx is not None:
        street_lines = block[1:city_state_zip_idx]
    else:
        # fallback: treat remaining lines as street
        street_lines = block[1:]

    street_address = clean(", ".join([l for l in street_lines if l]))

    return (company, street_address, city, state, zip_code)

def extract_pdf_text(pdf_path: Path) -> str:
    # Most Voelkr PDFs are text-based; if some are scanned images, this won’t work without OCR.
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip():
                parts.append(t)
    return "\n".join(parts)


# ---------- main extraction ----------
def extract_fields_from_pdf(pdf_path: Path) -> dict:
    text = extract_pdf_text(pdf_path)

    company, ship_addr, city, state, zip_code = extract_ship_to(text)

    # Label-based fields (tweak labels if your PDFs use different wording)
    salesperson = find_labeled_value(text, "Salesperson")
    quoted_by = find_labeled_value(text, "Quoted By")

    # Customer number label variants
    cust_num = (
        find_labeled_value(text, "Cust #")
        or find_labeled_value(text, "Cust#")
        or find_labeled_value(text, "Customer #")
        or find_labeled_value(text, "Customer#")
    )

    return {
        "File": pdf_path.name,
        "Company": company,
        "Ship To Address": ship_addr,
        "City": city,
        "State": state,
        "Zip": zip_code,
        "Salesperson": salesperson,
        "Quoted By": quoted_by,
        "Cust #": cust_num,
    }

def run_batch(input_folder: str, output_csv: str):
    input_path = Path(input_folder)
    pdfs = sorted(input_path.glob("*.pdf"))

    rows = []
    for pdf in pdfs:
        try:
            rows.append(extract_fields_from_pdf(pdf))
        except Exception as e:
            rows.append({
                "File": pdf.name,
                "Company": "",
                "Ship To Address": "",
                "City": "",
                "State": "",
                "Zip": "",
                "Salesperson": "",
                "Quoted By": "",
                "Cust #": "",
                "Error": str(e),
            })

    fieldnames = [
        "File", "Company", "Ship To Address", "City", "State", "Zip",
        "Salesperson", "Quoted By", "Cust #", "Error"
    ]
    # Ensure all keys exist
    for r in rows:
        for f in fieldnames:
            r.setdefault(f, "")

    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {output_csv}")


if __name__ == "__main__":
    # Example usage:
    #   put PDFs in ./voelkr_pdfs
    #   python extract_voelkr.py
    run_batch("./voelkr_pdfs", "voelkr_extracted.csv")
