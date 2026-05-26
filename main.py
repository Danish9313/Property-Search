"""
Debt Recovery & Collateralization Analysis Pipeline
====================================================
Reads an Excel/ODS file, matches property addresses to local JSON files,
builds a structured prompt per customer, calls the Claude API for analysis,
and writes results back into the spreadsheet.

Usage:
    python pipeline.py --excel path/to/file.ods --json_dir path/to/json_data/ [--output path/to/output.xlsx]

Requirements:
    pip install pandas openpyxl odfpy google-generativeai rapidfuzz
"""

import os
import re
import sys
import json
import argparse
import logging
from pathlib import Path
from typing import Optional
import pandas as pd
from rapidfuzz import process, fuzz
import config
def close_browser(): pass  # no-op — browser not used when calling API directly

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Homestead exemption lookup (per-state, USD)
# Source: National Consumer Law Center / individual state statutes (2024)
# ---------------------------------------------------------------------------
HOMESTEAD_EXEMPTIONS: dict[str, float] = {
    "AL": 15000, "AK": 54000, "AZ": 150000, "AR": 2500, "CA": 626400,
    "CO": 250000, "CT": 75000, "DE": 125000, "FL": 0,       # FL = unlimited
    "GA": 21500,  "HI": 30000, "ID": 175000, "IL": 15000,  "IN": 19300,
    "IA": 0,      # IA = unlimited
    "KS": 0,      # KS = unlimited
    "KY": 5000,   "LA": 35000, "ME": 80000, "MD": 25150, "MA": 500000,
    "MI": 40475,  "MN": 480000,"MS": 75000, "MO": 15000, "MT": 350000,
    "NE": 60000,  "NV": 605000,"NH": 120000,"NJ": 0,    "NM": 60000,
    "NY": 179950, "NC": 35000, "ND": 100000,"OH": 145425,"OK": 0,  # OK = unlimited
    "OR": 40000,  "PA": 0,     # PA = no homestead exemption
    "RI": 500000, "SC": 63075, "SD": 0,     # SD = unlimited
    "TN": 5000,   "TX": 0,     # TX = unlimited
    "UT": 42700,  "VT": 125000,"VA": 25000, "WA": 125000,"WV": 35000,
    "WI": 75000,  "WY": 20000, "DC": 0,
}

# States with "unlimited" exemption -- debt is not recoverable via primary residence
UNLIMITED_HOMESTEAD_STATES = {"FL", "IA", "KS", "OK", "SD", "TX"}

# ---------------------------------------------------------------------------
# SSN-sensitive columns to never include in prompts
# ---------------------------------------------------------------------------
BLOCKED_COLUMNS = {
    "SSN", "CUST_SSN", "CUST_SSN.1",
}

# ---------------------------------------------------------------------------
# Address normalization helpers
# ---------------------------------------------------------------------------
_STREET_ABBREVS = {
    r"\bST\b": "STREET", r"\bAVE\b": "AVENUE", r"\bBLVD\b": "BOULEVARD",
    r"\bDR\b": "DRIVE",  r"\bCT\b": "COURT",   r"\bRD\b": "ROAD",
    r"\bLN\b": "LANE",   r"\bPL\b": "PLACE",   r"\bPKWY\b": "PARKWAY",
    r"\bHWY\b": "HIGHWAY",r"\bCIR\b": "CIRCLE", r"\bTRL\b": "TRAIL",
    r"\bWAY\b": "WAY",
}

def normalize_address(addr: str) -> str:
    """Uppercase, strip punctuation, collapse whitespace, expand common abbreviations."""
    if not addr or not isinstance(addr, str):
        return ""
    addr = addr.upper().strip()
    addr = re.sub(r"[,#\.\-]", " ", addr)
    for abbrev, full in _STREET_ABBREVS.items():
        addr = re.sub(abbrev, full, addr)
    return re.sub(r"\s+", " ", addr).strip()


def json_filename_from_address(addr: str) -> str:
    """Convert an address string to the expected JSON filename pattern."""
    addr = addr.upper().strip()
    # Replace special chars with underscores (matches naming in JSON folder)
    safe = re.sub(r"[^A-Z0-9]", "_", addr)
    safe = re.sub(r"_+", "_", safe).strip("_")
    return safe + ".json"


# ---------------------------------------------------------------------------
# Customer data store (customer_id based, all files: TXT + PDF)
# ---------------------------------------------------------------------------
class CustomerDataStore:
    """
    Lazy-loading store: scans folder names at startup, reads files only when find() is called.
    Folder naming convention: "{customer_id} - {business_name}"
    The numeric prefix is used as the customer_id key.
    """
    def __init__(self, base_dir: str):
        self.base_dir = Path(base_dir)
        # Maps customer_id -> Path of their folder (populated at startup, fast)
        self._dirs: dict[str, Path] = {}
        self._scan_dirs()

    def _scan_dirs(self):
        """Scan only folder names — no file reading yet."""
        for customer_dir in self.base_dir.rglob("*"):
            if not customer_dir.is_dir():
                continue
            m = re.match(r'^(\d+)', customer_dir.name)
            if not m:
                continue
            self._dirs[m.group(1)] = customer_dir
        logger.info(f"Found folders for {len(self._dirs)} customers.")

    @staticmethod
    def _long_path(path: Path) -> str:
        """Extended-length path prefix on Windows to bypass MAX_PATH (260 char) limit."""
        if sys.platform == "win32":
            return "\\\\?\\" + str(path.resolve())
        return str(path.resolve())

    _PDF_MAX_PAGES = 20  # read at most 20 pages per PDF

    @staticmethod
    def _extract_pdf(path: Path) -> str:
        try:
            import pdfplumber
            text_parts = []
            with open(CustomerDataStore._long_path(path), "rb") as f:
                with pdfplumber.open(f) as pdf:
                    for page in pdf.pages[:CustomerDataStore._PDF_MAX_PAGES]:
                        text = page.extract_text()
                        if text:
                            text_parts.append(text)
            return "\n".join(text_parts).strip()
        except Exception as e:
            logger.warning(f"PDF extract failed {path.name}: {e}")
            return ""

    def find(self, customer_id) -> list[dict]:
        """Load and return all TXT + PDF files for the given customer_id on demand."""
        if customer_id is None:
            return []
        cid = str(customer_id).strip()
        customer_dir = self._dirs.get(cid)
        if customer_dir is None:
            return []

        data_list = []
        for path in sorted(customer_dir.glob("*.txt")):
            try:
                with open(self._long_path(path), encoding="utf-8", errors="replace") as f:
                    content = f.read().strip()
                if content:
                    data_list.append({"filename": path.name, "content": content})
            except Exception as e:
                logger.warning(f"Failed to load TXT {path.name}: {e}")

        PDF_SIZE_LIMIT_MB = 5
        for path in sorted(customer_dir.glob("*.pdf")):
            size_mb = path.stat().st_size / (1024 * 1024)
            if size_mb > PDF_SIZE_LIMIT_MB:
                logger.warning(f"  Skipping oversized PDF ({size_mb:.1f} MB): {path.name}")
                continue
            content = self._extract_pdf(path)
            if content:
                data_list.append({"filename": path.name, "content": content})

        return data_list


# ---------------------------------------------------------------------------
# JSON data extractor -> structured dict for the prompt
# ---------------------------------------------------------------------------
def extract_property_summary(json_data: dict, address_label: str) -> dict:
    """Pull key fields from a property JSON into a flat summary dict."""
    prop_results = (json_data.get("property") or {}).get("results") or []
    owners_results = (json_data.get("owners") or {}).get("results") or []
    transactions = (json_data.get("transactions") or {}).get("results") or []

    summary: dict = {"address_label": address_label}

    if prop_results:
        p = prop_results[0]
        summary["full_address"] = f"{p.get('Address','')}, {p.get('City','')}, {p.get('State','')} {p.get('ZipFive','')}"
        summary["state"] = p.get("State", "")
        summary["avm"] = p.get("AVM")
        summary["available_equity"] = p.get("AvailableEquity")
        summary["property_type"] = p.get("AdvancedPropertyType", p.get("PType"))
        summary["lot_size_sqft"] = p.get("LotSize")
        summary["is_same_mailing"] = bool(p.get("isSameMailingOrExempt"))
        summary["in_foreclosure"] = bool(p.get("inForeclosure"))
        summary["in_tax_delinquency"] = bool(p.get("inTaxDelinquency"))
        summary["in_bankruptcy"] = bool(p.get("inBankruptcyProperty"))
        summary["in_probate"] = bool(p.get("inProbateProperty"))
        summary["has_open_person_liens"] = bool(p.get("PropertyHasOpenPersonLiens"))
        summary["has_open_liens"] = bool(p.get("PropertyHasOpenLiens"))
        summary["last_transfer_date"] = p.get("LastTransferRecDate")
        summary["last_transfer_value"] = p.get("LastTransferValue")
        summary["distress_score"] = p.get("DistressScore")

    if owners_results:
        o = owners_results[0]
        summary["owner_name"] = f"{o.get('FirstName','')} {o.get('LastName','')}".strip()
        summary["owner_age"] = o.get("Age")
        summary["primary_residence_match"] = bool(o.get("PrimaryResidence"))

    # Extract active loans (not released) from transactions
    active_loans: list[dict] = []
    released_doc_numbers: set = set()

    # First pass: collect released loan doc numbers
    for t in transactions:
        if "release" in (t.get("DocTypeUI") or "").lower():
            released_doc_numbers.add(t.get("DocNumber"))

    # Second pass: collect loans not yet released
    for t in transactions:
        if t.get("DocTypeUI") == "Loan" and t.get("DocNumber") not in released_doc_numbers:
            active_loans.append({
                "position": t.get("LoanPosition"),
                "amount": t.get("Amount"),
                "grantor": t.get("Grantor"),
                "grantee": t.get("Grantee"),
                "date": t.get("RecDate"),
                "purpose": t.get("Purpose"),
            })

    summary["active_loans"] = active_loans
    summary["total_active_loan_amount"] = sum(
        (loan.get("amount") or 0) for loan in active_loans
    )

    # Comps summary for market context
    comps = (json_data.get("comps_sales") or {}).get("results") or []
    if comps:
        comp_values = [c.get("TransferValue", 0) for c in comps if c.get("TransferValue")]
        if comp_values:
            summary["comp_avg_value"] = round(sum(comp_values) / len(comp_values), 2)
            summary["comp_min_value"] = min(comp_values)
            summary["comp_max_value"] = max(comp_values)
            summary["comp_count"] = len(comp_values)

    return summary


# ---------------------------------------------------------------------------
# Homestead exemption resolver
# ---------------------------------------------------------------------------
def get_homestead_exemption(state: str, is_primary_residence: bool) -> tuple[float, str]:
    """
    Returns (exemption_amount, note).
    Unlimited states -> returns AVM value as exemption (handled downstream).
    PA -> 0 (no exemption).
    """
    if not is_primary_residence:
        return 0.0, "Not primary residence - no homestead exemption applied."

    state = (state or "").upper().strip()
    if state in UNLIMITED_HOMESTEAD_STATES:
        return -1.0, f"{state} has unlimited homestead exemption - primary residence may be fully protected."
    amount = HOMESTEAD_EXEMPTIONS.get(state, 0.0)
    if amount == 0.0:
        return 0.0, f"{state} has no homestead exemption (or $0 applicable)."
    return amount, f"{state} homestead exemption: ${amount:,.0f} applied to primary residence."


# ---------------------------------------------------------------------------
# Prompt builder
# ---------------------------------------------------------------------------
PROMPT_TEMPLATE_PATH = Path(__file__).parent / "prompt_template.txt"

def _load_prompt_template() -> str:
    """Load the prompt template from prompt_template.txt."""
    try:
        return PROMPT_TEMPLATE_PATH.read_text(encoding="utf-8")
    except FileNotFoundError:
        raise FileNotFoundError(f"prompt_template.txt not found at {PROMPT_TEMPLATE_PATH}")


def build_analysis_prompt(row: pd.Series, property_data: list[tuple]) -> str:
    """
    Build the analysis prompt by loading prompt_template.txt and injecting
    the Excel row data, property JSON summaries, and output schema.
    """
    import math

    # Build filtered row context (exclude blocked + estimated equity columns)
    row_context: dict = {}
    for col, val in row.items():
        if col in BLOCKED_COLUMNS:
            continue
        if "estimated equity" in str(col).lower():
            continue
        if val is None or (isinstance(val, float) and math.isnan(val)):
            continue
        row_context[str(col)] = str(val).strip()

    # Build per-property sections from raw TXT content
    props_payload = []
    missing_json = []
    for label, filename, content in property_data:
        if content:
            props_payload.append({"label": label, "filename": filename, "content": content})
        else:
            missing_json.append(f"{label} ({filename})")

    output_schema = {
        "skip": False,
        "Deceased Check": "",
        "Business Bankruptcy": "",
        "PG Bankruptcy Check": "",
        "SOL": "",
        "SOL Status": "",
        "Verified Equity": "",
        "Verified Liens": "",
        "Homestead Applied": "",
        "Homestead State": "",
        "Final Collateral": "",
        "Collateral Calculation": "",
        "Collateralization": "",
        "Principal Balance": "",
        "Collectibility Judgment": "",
        "Recovery Summary": "",
        "Criminal records:": "",
        "Other Assets:": "",
        "Professional Licenses:": "",
        "Other Owned Businesses:": "",
        "notes": "",
    }

    missing_note = (
        f"NOTE: No JSON data found for: {', '.join(missing_json)}. Those properties were excluded."
        if missing_json else ""
    )

    template = _load_prompt_template()
    prompt = template.replace("{EXCEL_ROW_DATA}", json.dumps(row_context, indent=2))
    prompt = prompt.replace("{PROPERTY_DATA}", json.dumps(props_payload, indent=2))
    prompt = prompt.replace("{MISSING_NOTE}", missing_note)
    prompt = prompt.replace("{OUTPUT_SCHEMA}", json.dumps(output_schema, indent=2))
    return prompt.strip()


# ---------------------------------------------------------------------------
# OpenRouter API caller
# ---------------------------------------------------------------------------
def call_claude_api(prompt: str, api_key: str) -> str:
    """Submit the analysis prompt to OpenRouter and return the response."""
    from openrouter import OpenRouter
    with OpenRouter(api_key=api_key, timeout_ms=120000) as client:
        response = client.chat.send(
            model=config.OPENROUTER_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "You are a financial analysis assistant. Always respond with ONLY a valid JSON object. No preamble, no explanation, no markdown. Start with { and end with }."
                },
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            max_tokens=config.OPENROUTER_MAX_OUTPUT_TOKENS,
            temperature=0.1,
        )
    content = response.choices[0].message.content
    return str(content) if content is not None else ""


# ---------------------------------------------------------------------------
# Response parser
# ---------------------------------------------------------------------------
def parse_ai_response(response: str) -> dict:
    """
    Parse the AI JSON response into a dict matching AI_OUTPUT_COLUMNS.
    Falls back to empty values if JSON is malformed.
    """
    blank = {col: "" for col in AI_OUTPUT_COLUMNS}

    # Write raw response to debug file for inspection
    try:
        with open("debug_last_response.txt", "w", encoding="utf-8") as f:
            f.write(response)
    except Exception:
        pass

    # Strip markdown code fences if present
    text = response.strip()
    text = re.sub(r"```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = text.strip()

    def _try_parse(s: str):
        """Attempt json.loads; if it fails due to unescaped newlines inside strings,
        replace literal newlines within quoted values and retry."""
        try:
            return json.loads(s)
        except json.JSONDecodeError:
            # Replace literal newlines inside JSON string values with \\n
            fixed = re.sub(r'(?<=": ")([^"]*)\n([^"]*)', r'\1\\n\2', s)
            fixed = re.sub(r'(?<=": ")([^"]*)\n([^"]*)', r'\1\\n\2', fixed)
            return json.loads(fixed)

    try:
        parsed = _try_parse(text)
    except json.JSONDecodeError:
        # Try to extract JSON object from surrounding text
        m = re.search(r"\{[\s\S]+\}", text)
        if m:
            try:
                parsed = _try_parse(m.group(0))
            except json.JSONDecodeError:
                logger.warning("Could not parse AI response as JSON.")
                blank["notes"] = f"JSON parse error. Raw: {response[:300]}"
                return blank
        else:
            logger.warning("No JSON object found in AI response.")
            blank["notes"] = f"No JSON found. Raw: {response[:300]}"
            return blank

    def _get(key):
        val = parsed.get(key)
        if val is None and key.endswith(":"):
            val = parsed.get(key[:-1])  # try without trailing colon
        return "" if val is None else str(val)

    return {
        "Deceased Check":            _get("Deceased Check"),
        "Business Bankruptcy":       _get("Business Bankruptcy"),
        "PG Bankruptcy Check":       _get("PG Bankruptcy Check"),
        "SOL":                       _get("SOL"),
        "SOL Status":                _get("SOL Status"),
        "Verified Equity":           _get("Verified Equity"),
        "Verified Liens":            _get("Verified Liens"),
        "Homestead Applied":         _get("Homestead Applied"),
        "Homestead State":           _get("Homestead State"),
        "Final Collateral":          _get("Final Collateral"),
        "Collateral Calculation":    _get("Collateral Calculation"),
        "Collateralization":         _get("Collateralization"),
        "Principal Balance":         _get("Principal Balance"),
        "Collectibility Judgment":   _get("Collectibility Judgment"),
        "Recovery Summary":          _get("Recovery Summary"),
        "Criminal records:":         _get("Criminal records:"),
        "Other Assets:":             _get("Other Assets:"),
        "Professional Licenses:":    _get("Professional Licenses:"),
        "Other Owned Businesses:":   _get("Other Owned Businesses:"),
        "notes":                     _get("notes"),
    }


# ---------------------------------------------------------------------------
# Row processor
# ---------------------------------------------------------------------------
def detect_property_columns(columns: list[str]) -> list[tuple[str, str, str]]:
    """
    Dynamically detect all property address columns and their matching
    Estimated Equity columns from the Excel headers.

    Matches any column whose name (case-insensitive) contains:
      - "owned property" (e.g. "Owner 1 owned property1", "Owner 2 owned property3")
      - "guarantor owned property" (e.g. "Guarantor owned property 1")

    For each address column found, looks for a matching equity column that:
      - Starts with the same base text AND contains "estimated equity"

    Returns list of (address_col, equity_col_or_empty, label)
    """
    results = []
    cols_lower = {c.lower(): c for c in columns}

    for col in columns:
        col_lower = col.lower().strip()

        # Skip equity columns themselves
        if "estimated equity" in col_lower:
            continue

        # Match owner/guarantor property address columns
        is_property_col = (
            "owned property" in col_lower or
            "guarantor owned property" in col_lower
        )
        if not is_property_col:
            continue

        # Build a human-readable label
        label = col.strip()

        # Find matching equity column: same prefix + "estimated equity"
        equity_col = ""
        col_base = col_lower.replace(" ", "")
        for c_low, c_orig in cols_lower.items():
            if "estimated equity" in c_low:
                # Check if the equity col starts with same base (ignoring spaces)
                equity_base = c_low.replace(" ", "").replace("estimatedequity$", "").replace("estimatedequity", "")
                if col_base.startswith(equity_base) or equity_base.startswith(col_base):
                    equity_col = c_orig
                    break

        results.append((col, equity_col, label))

    return results


def process_row(
    row: pd.Series,
    json_store: CustomerDataStore,
    api_key: str,
    id_col: str = "customer_id",
) -> dict:
    """
    Process a single Excel row:
    1. Filter: skip if Property Status != "property found"
    2. Look up all files by id_col folder match
    3. Build prompt with full row + file data
    4. Call Gemini and parse JSON response
    """
    cust_id = str(row.get(id_col) or row.get("customer_id") or row.get("CUST_NUMBER") or "?").strip()

    # ROW FILTER — accept "property status" or "property details" column
    # Look up all files for this row by id_col folder match
    txt_files = json_store.find(cust_id)

    if not txt_files:
        logger.warning(f"  Row {cust_id}: No property folder found for '{id_col}' = '{cust_id}'.")
        blank = {col: "" for col in AI_OUTPUT_COLUMNS}
        blank["Collectibility Judgment"] = "Low"
        blank["notes"] = f"No property folder found for '{id_col}' = '{cust_id}'."
        return blank

    # Build property_data from all loaded files
    property_data: list[tuple] = []
    for i, item in enumerate(txt_files, 1):
        label = f"Property {i}"
        filename = item["filename"]
        content = item["content"]
        property_data.append((label, filename, content))
        logger.info(f"  Row {cust_id}: Loaded [{label}] -> '{filename}'")

    # Use manually-filled Verified Equity if present, otherwise let AI determine it
    manual_ve = str(row.get("Verified Equity", "") or "").strip()
    if manual_ve and manual_ve.lower() not in ("nan", ""):
        try:
            ve_num = float(manual_ve.replace("$", "").replace(",", ""))
            verified_equity = f"${ve_num:,.2f}"
        except (ValueError, TypeError):
            verified_equity = manual_ve
    else:
        verified_equity = None  # None = let AI value through

    prompt = build_analysis_prompt(row, property_data)
    logger.info(f"  Row {cust_id}: Calling OpenRouter API...")
    ai_response = call_claude_api(prompt, api_key)
    result = parse_ai_response(ai_response)

    # Only override Verified Equity if manual value exists
    if verified_equity is not None:
        result["Verified Equity"] = verified_equity

    # Compute Collateralization from Final Collateral
    try:
        fc = float(str(result.get("Final Collateral", "0")).replace("$", "").replace(",", "").strip())
        result["Collateralization"] = "Positive Collateral" if fc > 0 else "Negative Collateral"
    except (ValueError, TypeError):
        result["Collateralization"] = ""

    return result


# ---------------------------------------------------------------------------
# Output writer
# ---------------------------------------------------------------------------
AI_OUTPUT_COLUMNS = [
    "Deceased Check",
    "Business Bankruptcy",
    "PG Bankruptcy Check",
    "SOL",
    "SOL Status",
    "Verified Equity",
    "Verified Liens",
    "Homestead Applied",
    "Homestead State",
    "Final Collateral",
    "Collateral Calculation",
    "Collateralization",
    "Principal Balance",
    "Collectibility Judgment",
    "Recovery Summary",
    "Criminal records:",
    "Other Assets:",
    "Professional Licenses:",
    "Other Owned Businesses:",
    "notes",
]

def write_results(
    original_df: pd.DataFrame,
    results: list[dict],
    output_path: str,
):
    """
    Write results back to Excel:
    - Sheet 1 (DataSheet): original data + AI result columns appended
    - Sheet 2 (AI_Analysis): clean summary view
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    # Build results dataframe
    results_df = pd.DataFrame(results, columns=AI_OUTPUT_COLUMNS)

    # Drop SSN cols and any columns that the AI output will overwrite
    drop_cols = list(BLOCKED_COLUMNS) + [c for c in AI_OUTPUT_COLUMNS if c in original_df.columns]
    safe_df = original_df.drop(columns=[c for c in drop_cols if c in original_df.columns], errors="ignore")
    combined_df = pd.concat([safe_df.reset_index(drop=True), results_df.reset_index(drop=True)], axis=1)

    # Write to Excel — if file is locked (open in Excel), save with timestamp
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            combined_df.to_excel(writer, sheet_name="DataSheet", index=False)
            results_df.to_excel(writer, sheet_name="AI_Analysis", index=False)
    except PermissionError:
        from datetime import datetime
        stem = Path(output_path).stem
        parent = Path(output_path).parent
        output_path = str(parent / f"{stem}_{datetime.now().strftime('%H%M%S')}.xlsx")
        logger.warning(f"Output file locked — saving to: {output_path}")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            combined_df.to_excel(writer, sheet_name="DataSheet", index=False)
            results_df.to_excel(writer, sheet_name="AI_Analysis", index=False)

    # Apply formatting
    from openpyxl import load_workbook
    wb = load_workbook(output_path)

    for sheet_name in ["DataSheet", "AI_Analysis"]:
        ws = wb[sheet_name]
        # Header row: bold, light blue fill
        header_fill = PatternFill("solid", start_color="BDD7EE")
        for cell in ws[1]:
            cell.font = Font(bold=True, name="Arial", size=10)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

        # Highlight AI result columns in DataSheet
        if sheet_name == "DataSheet":
            ai_col_names = set(AI_OUTPUT_COLUMNS)
            for col_idx, cell in enumerate(ws[1], 1):
                if cell.value in ai_col_names:
                    ai_fill = PatternFill("solid", start_color="E2EFDA")
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                        for c in row:
                            c.fill = ai_fill

        # Color-code collectibility
        if sheet_name == "AI_Analysis":
            col_headers = {cell.value: cell.column for cell in ws[1]}
            collectibility_col = col_headers.get("Collectibility Judgment")
            if collectibility_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                        min_col=collectibility_col, max_col=collectibility_col):
                    for cell in row:
                        val = str(cell.value or "").lower()
                        if "high" in val:
                            cell.fill = PatternFill("solid", start_color="C6EFCE")
                            cell.font = Font(color="375623", name="Arial")
                        elif "medium" in val:
                            cell.fill = PatternFill("solid", start_color="FFEB9C")
                            cell.font = Font(color="9C5700", name="Arial")
                        elif "low" in val:
                            cell.fill = PatternFill("solid", start_color="FFC7CE")
                            cell.font = Font(color="9C0006", name="Arial")

        # Auto-width columns (capped at 60)
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

        # Wrap text for long AI columns
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)
    logger.info(f"Output written to: {output_path}")


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Debt Recovery Analysis Pipeline -- all defaults loaded from config.py / .env"
    )
    # All args are optional -- config.py provides the defaults
    parser.add_argument("--excel",   default=None, help=f"Excel/ODS file path (default: {config.EXCEL_FILE_PATH})")
    parser.add_argument("--json_dir",default=None, help=f"JSON folder path (default: {config.JSON_FOLDER_PATH})")
    parser.add_argument("--output",  default=None, help=f"Output Excel path (default: {config.OUTPUT_FILE_PATH})")
    parser.add_argument("--api_key", default=None, help="OpenRouter API key (default: OPENROUTER_API_KEY from .env)")
    parser.add_argument("--max_rows",type=int, default=None, help="Limit rows for testing (default: all rows)")
    parser.add_argument("--start_from", type=int, default=0, help="Skip first N rows (resume from row N+1)")
    parser.add_argument("--id_col",  default="customer_id", help="Column name to match with folder prefix (default: customer_id)")
    parser.add_argument("--sheet",   default=None, help="Excel sheet name to read (default: first sheet)")
    args = parser.parse_args()

    # CLI args override config.py, config.py overrides defaults
    excel_path   = args.excel    or config.EXCEL_FILE_PATH
    json_dir     = args.json_dir or config.JSON_FOLDER_PATH
    output_path  = args.output   or config.OUTPUT_FILE_PATH
    api_key      = args.api_key  or config.OPENROUTER_API_KEY
    max_rows     = args.max_rows if args.max_rows is not None else (config.MAX_ROWS or None)
    id_col       = args.id_col
    start_from   = args.start_from
    sheet_name   = args.sheet

    # Validate only API key; skip path validation since CLI args may override .env paths
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY is not set. Add it to your .env file.")
    if not Path(excel_path).exists():
        raise ValueError(f"Excel file not found: '{excel_path}'")
    if not Path(json_dir).exists():
        raise ValueError(f"JSON folder not found: '{json_dir}'")

    logger.info("Pipeline starting with config:")
    logger.info(f"  Excel file  : {excel_path}")
    logger.info(f"  JSON folder : {json_dir}")
    logger.info(f"  Output file : {output_path}")
    logger.info(f"  OpenRouter model: {config.OPENROUTER_MODEL}")
    logger.info(f"  Max rows    : {max_rows if max_rows else 'All'}")
    logger.info(f"  ID column   : {id_col}")

    # Load Excel
    ext = Path(excel_path).suffix.lower()
    if ext == ".ods":
        df = pd.read_excel(excel_path, engine="odf", sheet_name=sheet_name or 0, dtype=str)
    else:
        df = pd.read_excel(excel_path, sheet_name=sheet_name or 0, dtype=str)

    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns.")

    full_df = df.copy()  # keep full dataframe for writing results
    if start_from > 0:
        df = df.iloc[start_from:].reset_index(drop=True)
        logger.info(f"Resuming from row {start_from + 1} (skipping first {start_from} rows).")
    if max_rows:
        df = df.head(max_rows)
        logger.info(f"Processing limited to {max_rows} rows.")

    # Load customer data store
    json_store = CustomerDataStore(json_dir)

    # Process each row — save after every row so results are visible in real time
    blank_result = {col: "" for col in AI_OUTPUT_COLUMNS}

    # If resuming, load existing results from output file for already-processed rows
    all_results: list[dict] = []
    if start_from > 0 and Path(output_path).exists():
        try:
            existing = pd.read_excel(output_path, sheet_name="AI_Analysis", dtype=str)
            for _, erow in existing.iterrows():
                all_results.append({col: str(erow.get(col, "") or "") for col in AI_OUTPUT_COLUMNS})
            logger.info(f"Loaded {len(all_results)} existing results from {output_path}")
        except Exception as e:
            logger.warning(f"Could not load existing results: {e}")
            all_results = [dict(blank_result) for _ in range(start_from)]
    else:
        all_results = []

    # Fill remaining slots with blanks
    total_rows_in_excel = start_from + len(df)
    while len(all_results) < total_rows_in_excel:
        all_results.append(dict(blank_result))

    for idx, row in df.iterrows():
        row_id = str(row.get(id_col) or row.get('customer_id') or row.get('CUST_NUMBER') or '?').strip()
        logger.info(f"Processing row {idx + 1}/{len(df)} | {id_col}={row_id}")
        try:
            result = process_row(row, json_store, api_key, id_col=id_col)
            if result.get("_skip"):
                result = {col: "" for col in AI_OUTPUT_COLUMNS}
                result["notes"] = "Skipped — Property Status != 'property found'"
        except Exception as e:
            logger.error(f"  Row {idx}: ERROR - {e}")
            result = {col: "" for col in AI_OUTPUT_COLUMNS}
            result["notes"] = f"PIPELINE ERROR: {e}"
        all_results[start_from + idx] = result
        # Only save after AI-processed rows (not skipped rows) to avoid slow Excel writes
        was_skipped = result.get("notes", "").startswith("Skipped")
        if not was_skipped:
            try:
                write_results(full_df, all_results, output_path)
                logger.info(f"  Saved progress ({idx + 1}/{len(df)} rows).")
            except Exception as e:
                logger.warning(f"  Could not save progress: {e}")

    # Final write — ensures skipped/trailing rows are included in AI_Analysis
    try:
        write_results(full_df, all_results, output_path)
        logger.info(f"Final output written to: {output_path}")
    except Exception as e:
        logger.warning(f"Could not write final output: {e}")

    close_browser()
    logger.info("Pipeline complete.")


if __name__ == "__main__":
    main()