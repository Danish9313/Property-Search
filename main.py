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
import json
import argparse
import logging
from pathlib import Path
from typing import Optional
import pandas as pd
from rapidfuzz import process, fuzz
import config
from browser_gemini import call_gemini_browser, close_browser

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
# JSON folder loader
# ---------------------------------------------------------------------------
class JsonPropertyStore:
    """
    Loads all JSON files from a directory once at startup.
    Keys: normalized address string  -> parsed JSON dict
    """
    def __init__(self, json_dir: str):
        self.json_dir = Path(json_dir)
        # Map: normalized_address -> (raw_address_key, json_data)
        self._store: dict[str, tuple[str, dict]] = {}
        self._load_all()

    def _load_all(self):
        for path in self.json_dir.glob("*.json"):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # The property address lives in property.results[0].Address + City + State + Zip
                prop_result = (data.get("property") or {}).get("results") or []
                if prop_result:
                    r = prop_result[0]
                    raw_addr = f"{r.get('Address','')}, {r.get('City','')}, {r.get('State','')} {r.get('ZipFive','')}"
                else:
                    # Fallback: derive from filename
                    raw_addr = path.stem.replace("_", " ")
                norm = normalize_address(raw_addr)
                self._store[norm] = (raw_addr, data)
            except Exception as e:
                logger.warning(f"Failed to load JSON {path.name}: {e}")
        logger.info(f"Loaded {len(self._store)} JSON property files.")

    def find(self, address: str, score_cutoff: int = None) -> Optional[dict]:
        """
        Fuzzy-match the given address against all loaded JSON keys.
        Returns the JSON dict if a match is found above score_cutoff, else None.
        """
        if score_cutoff is None:
            score_cutoff = config.ADDRESS_MATCH_SCORE_CUTOFF
        if not address or not isinstance(address, str):
            return None
        norm = normalize_address(address)
        if not norm:
            return None

        keys = list(self._store.keys())
        result = process.extractOne(norm, keys, scorer=fuzz.token_sort_ratio, score_cutoff=score_cutoff)
        if result:
            matched_key, score, _ = result
            raw_addr, data = self._store[matched_key]
            logger.debug(f"Match: '{address}' -> '{raw_addr}' (score={score})")
            return data
        return None


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
SYSTEM_PROMPT = """You are an AI assistant processing property data for debt recovery analysis.
Analyze the provided Excel row and property JSON data to determine debt collectibility.
Always respond with valid JSON only — no preamble, no explanation, no markdown code fences.

Use the following homestead exemption table (2024) when calculating net collateral:
AL=$15,000 | AK=$54,000 | AZ=$150,000 | AR=$2,500 | CA=$626,400 | CO=$250,000
CT=$75,000 | DE=$125,000 | FL=UNLIMITED | GA=$21,500 | HI=$30,000 | ID=$175,000
IL=$15,000 | IN=$19,300 | IA=UNLIMITED | KS=UNLIMITED | KY=$5,000 | LA=$35,000
ME=$80,000 | MD=$25,150 | MA=$500,000 | MI=$40,475 | MN=$480,000 | MS=$75,000
MO=$15,000 | MT=$350,000 | NE=$60,000 | NV=$605,000 | NH=$120,000 | NJ=$0
NM=$60,000 | NY=$179,950 | NC=$35,000 | ND=$100,000 | OH=$145,425 | OK=UNLIMITED
OR=$40,000 | PA=$0 | RI=$500,000 | SC=$63,075 | SD=UNLIMITED | TN=$5,000
TX=UNLIMITED | UT=$42,700 | VT=$125,000 | VA=$25,000 | WA=$125,000 | WV=$35,000
WI=$75,000 | WY=$20,000 | DC=$0

UNLIMITED states (FL, IA, KS, OK, SD, TX): primary residence is fully protected — net
collateral from that property is $0 after homestead."""


def build_analysis_prompt(row: pd.Series, property_data: list[tuple]) -> str:
    """
    Build the analysis prompt combining the full Excel row with matched JSON files.
    property_data: list of (label, address, json_data_or_None) tuples.
    Returns a prompt requesting strict JSON output.
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

    # Build per-property sections
    props_payload = []
    missing_json = []
    for label, address, json_data in property_data:
        if json_data is not None:
            props_payload.append({"label": label, "address": address, "data": json_data})
        else:
            missing_json.append(f"{label} ({address})")

    output_schema = {
        "skip": False,
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
        "Lien Enforcement": "",
        "notes": "",
    }

    missing_note = (
        f"No JSON data found for: {', '.join(missing_json)}. Those properties were ignored."
        if missing_json else ""
    )

    prompt = f"""You are an AI assistant processing property data from an Excel file.

EXCEL ROW DATA:
{json.dumps(row_context, indent=2)}

PROPERTY JSON DATA:
{json.dumps(props_payload, indent=2)}

INSTRUCTIONS:
1. Use the FULL Excel row for context.
2. Analyze ALL properties together to produce FINAL financial/legal outputs for this row.
3. Ignore any "Estimated Equity $" columns from the Excel row.
4. Do NOT mix incorrect data across properties, but the final decision must consider all valid properties.
5. Apply homestead exemption ONLY to the primary residence (isSameMailingOrExempt=true or PrimaryResidence=true).
6. For Collectibility Judgment use: Low / Medium / High
   - High: Final Collateral >= Principal Balance
   - Medium: 0 < Final Collateral < Principal Balance
   - Low: Final Collateral <= 0
7. Use null (not empty string) for any field where data is genuinely unavailable.
{f'9. NOTE: {missing_note}' if missing_note else ''}

OUTPUT: Return ONLY valid JSON matching this exact schema — no explanation, no markdown:
{json.dumps(output_schema, indent=2)}
"""
    return prompt.strip()

# ---------------------------------------------------------------------------
# Gemini browser caller
# ---------------------------------------------------------------------------
def call_claude_api(prompt: str, api_key: str) -> str:
    """Submit the analysis prompt to Gemini via browser (web search enabled) and return the response."""
    return call_gemini_browser(prompt)


# ---------------------------------------------------------------------------
# Response parser
# ---------------------------------------------------------------------------
def parse_ai_response(response: str) -> dict:
    """
    Parse the AI JSON response into a dict matching AI_OUTPUT_COLUMNS.
    Falls back to empty values if JSON is malformed.
    """
    blank = {col: "" for col in AI_OUTPUT_COLUMNS}

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
        return "" if val is None else str(val)

    return {
        "Verified Equity":        _get("Verified Equity"),
        "Verified Liens":         _get("Verified Liens"),
        "Homestead Applied":      _get("Homestead Applied"),
        "Homestead State":        _get("Homestead State"),
        "Final Collateral":       _get("Final Collateral"),
        "Collateral Calculation": _get("Collateral Calculation"),
        "Collateralization":      _get("Collateralization"),
        "Principal Balance":      _get("Principal Balance"),
        "Collectibility Judgment":_get("Collectibility Judgment"),
        "Recovery Summary":       _get("Recovery Summary"),
        "Lien Enforcement":       _get("Lien Enforcement"),
        "notes":                  _get("notes"),
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
    json_store: JsonPropertyStore,
    api_key: str,
) -> dict:
    """
    Process a single Excel row:
    1. Filter: skip if Property Status != "property found"
    2. Find JSON files for each property address
    3. Build prompt with full row + JSON data
    4. Call Gemini and parse JSON response
    """
    cust_id = row.get("CUST_NUMBER", "?")

    # ROW FILTER — case-insensitive column lookup, value contains "property found"
    status_col = next((c for c in row.index if c.lower() == "property status"), None)
    property_status = str(row.get(status_col, "") or "").strip().lower() if status_col else ""
    if "property found" not in property_status:
        logger.info(f"  Row {cust_id}: Skipping — Property Status = '{property_status}'")
        return {"_skip": True}

    # Dynamically detect property columns from this row's index
    property_cols = detect_property_columns(list(row.index))

    # Collect property addresses and their JSON data
    property_data: list[tuple] = []
    for addr_col, _equity_col, label in property_cols:
        addr_val = row.get(addr_col)
        if not addr_val or not isinstance(addr_val, str) or not addr_val.strip():
            continue
        logger.info(f"  Row {cust_id}: Looking up JSON for [{label}] -> '{addr_val}'")
        json_data = json_store.find(addr_val)
        if not json_data:
            logger.warning(f"  Row {cust_id}: No JSON match for '{addr_val}'")
        property_data.append((label, addr_val, json_data))

    if not property_data:
        logger.warning(f"  Row {cust_id}: No property addresses found. Skipping AI call.")
        blank = {col: "" for col in AI_OUTPUT_COLUMNS}
        blank["Collectibility Judgment"] = "Low"
        blank["notes"] = "No property addresses found in this row."
        return blank

    # Calculate Verified Equity from Excel estimated equity columns
    equity_total = 0.0
    for _addr_col, equity_col, _label in property_cols:
        if not equity_col:
            continue
        val = row.get(equity_col)
        num = pd.to_numeric(str(val).replace(",", "").replace("$", "").strip(), errors="coerce")
        if num is not None and not pd.isna(num) and num > 0:
            equity_total += num
    verified_equity = f"${equity_total:,.2f}" if equity_total > 0 else ""

    prompt = build_analysis_prompt(row, property_data)
    logger.info(f"  Row {cust_id}: Calling Gemini API...")
    ai_response = call_claude_api(prompt, api_key)
    result = parse_ai_response(ai_response)

    # Override Verified Equity with Excel values
    result["Verified Equity"] = verified_equity

    result["Collateralization"] = "0"

    return result


# ---------------------------------------------------------------------------
# Output writer
# ---------------------------------------------------------------------------
AI_OUTPUT_COLUMNS = [
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
    "Lien Enforcement",
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

    # Write to Excel
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
    parser.add_argument("--api_key", default=None, help="Gemini API key (default: GEMINI_API_KEY from .env)")
    parser.add_argument("--max_rows",type=int, default=None, help="Limit rows for testing (default: all rows)")
    args = parser.parse_args()

    # CLI args override config.py, config.py overrides defaults
    excel_path   = args.excel    or config.EXCEL_FILE_PATH
    json_dir     = args.json_dir or config.JSON_FOLDER_PATH
    output_path  = args.output   or config.OUTPUT_FILE_PATH
    api_key      = args.api_key  or config.GEMINI_API_KEY
    max_rows     = args.max_rows if args.max_rows is not None else (config.MAX_ROWS or None)

    # Validate config before doing any work
    config.validate()

    logger.info("Pipeline starting with config:")
    logger.info(f"  Excel file  : {excel_path}")
    logger.info(f"  JSON folder : {json_dir}")
    logger.info(f"  Output file : {output_path}")
    logger.info(f"  Gemini model: {config.GEMINI_MODEL}")
    logger.info(f"  Max rows    : {max_rows if max_rows else 'All'}")

    # Load Excel
    ext = Path(excel_path).suffix.lower()
    if ext == ".ods":
        df = pd.read_excel(excel_path, engine="odf", dtype=str)
    else:
        df = pd.read_excel(excel_path, dtype=str)

    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns.")

    if max_rows:
        df = df.head(max_rows)
        logger.info(f"Processing limited to {max_rows} rows.")

    # Load JSON store
    json_store = JsonPropertyStore(json_dir)

    # Process each row
    all_results: list[dict] = []
    for idx, row in df.iterrows():
        logger.info(f"Processing row {idx + 1}/{len(df)} | CUST_NUMBER={row.get('CUST_NUMBER', '?')}")
        try:
            result = process_row(row, json_store, api_key)
            if result.get("_skip"):
                result = {col: "" for col in AI_OUTPUT_COLUMNS}
                result["notes"] = "Skipped — Property Status != 'property found'"
        except Exception as e:
            logger.error(f"  Row {idx}: ERROR - {e}")
            result = {col: "" for col in AI_OUTPUT_COLUMNS}
            result["notes"] = f"PIPELINE ERROR: {e}"
        all_results.append(result)

    # Write output
    write_results(df, all_results, output_path)
    close_browser()
    logger.info("Pipeline complete.")


if __name__ == "__main__":
    main()