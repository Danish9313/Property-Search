"""
config.py
=========
Central configuration for the Debt Recovery Analysis Pipeline.
All settings are loaded from the .env file or environment variables.
To override any value, update your .env file -- never hardcode keys here.
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# OpenRouter API
# ---------------------------------------------------------------------------

# Your OpenRouter API key from https://openrouter.ai/keys
OPENROUTER_API_KEY: str = os.getenv("OPENROUTER_API_KEY", "")

# OpenRouter model to use for analysis
# Options: "google/gemini-2.5-pro", "google/gemini-2.5-flash", "anthropic/claude-3-5-sonnet"
OPENROUTER_MODEL: str = os.getenv("OPENROUTER_MODEL", "google/gemini-2.5-pro")

# Max tokens in the AI response per customer row
OPENROUTER_MAX_OUTPUT_TOKENS: int = int(os.getenv("OPENROUTER_MAX_OUTPUT_TOKENS", "8192"))

# Keep these as aliases so any code referencing old names still works
GEMINI_API_KEY: str = OPENROUTER_API_KEY
GEMINI_MODEL: str = OPENROUTER_MODEL
GEMINI_MAX_OUTPUT_TOKENS: int = OPENROUTER_MAX_OUTPUT_TOKENS

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------

# Path to the input Excel or ODS file
EXCEL_FILE_PATH: str = os.getenv("EXCEL_FILE_PATH", "test_property_search.ods")

# Path to the folder containing property JSON files
JSON_FOLDER_PATH: str = os.getenv("JSON_FOLDER_PATH", "json_data")

# Output Excel file path
OUTPUT_FILE_PATH: str = os.getenv("OUTPUT_FILE_PATH", "debt_recovery_output.xlsx")

# ---------------------------------------------------------------------------
# Browser / Playwright settings
# ---------------------------------------------------------------------------

# Set to true to run the browser in headless mode (no visible window)
# Leave false so you can see Gemini working and log in on first run
BROWSER_HEADLESS: bool = os.getenv("BROWSER_HEADLESS", "false").lower() in ("true", "1", "yes")

# ---------------------------------------------------------------------------
# Pipeline behavior
# ---------------------------------------------------------------------------

# Limit number of rows processed (set to 0 for all rows)
MAX_ROWS: int = int(os.getenv("MAX_ROWS", "0"))

# Fuzzy match score cutoff for address matching (0-100)
# Lower = more lenient matching, higher = stricter
ADDRESS_MATCH_SCORE_CUTOFF: int = int(os.getenv("ADDRESS_MATCH_SCORE_CUTOFF", "72"))

# ---------------------------------------------------------------------------
# Validation -- called at startup to catch missing required config early
# ---------------------------------------------------------------------------

def validate():
    """Raise an error early if required config values are missing."""
    errors = []

    if not OPENROUTER_API_KEY:
        errors.append(
            "OPENROUTER_API_KEY is not set. Add it to your .env file:\n"
            "  OPENROUTER_API_KEY=sk-or-v1-..."
        )

    if not Path(EXCEL_FILE_PATH).exists():
        errors.append(
            f"EXCEL_FILE_PATH not found: '{EXCEL_FILE_PATH}'\n"
            "  Update EXCEL_FILE_PATH in your .env file."
        )

    if not Path(JSON_FOLDER_PATH).exists():
        errors.append(
            f"JSON_FOLDER_PATH not found: '{JSON_FOLDER_PATH}'\n"
            "  Update JSON_FOLDER_PATH in your .env file."
        )

    if errors:
        raise ValueError(
            "Config validation failed:\n\n" + "\n\n".join(f"- {e}" for e in errors)
        )


if __name__ == "__main__":
    # Quick check: print current config (masks API key)
    print("Current configuration:")
    print(f"  OPENROUTER_API_KEY    : {'*' * 8 + OPENROUTER_API_KEY[-4:] if OPENROUTER_API_KEY else 'NOT SET'}")
    print(f"  OPENROUTER_MODEL      : {OPENROUTER_MODEL}")
    print(f"  OPENROUTER_MAX_OUTPUT_TOKENS: {OPENROUTER_MAX_OUTPUT_TOKENS}")
    print(f"  EXCEL_FILE_PATH       : {EXCEL_FILE_PATH}")
    print(f"  JSON_FOLDER_PATH      : {JSON_FOLDER_PATH}")
    print(f"  OUTPUT_FILE_PATH      : {OUTPUT_FILE_PATH}")
    print(f"  MAX_ROWS              : {MAX_ROWS if MAX_ROWS > 0 else 'All rows'}")
    print(f"  ADDRESS_MATCH_SCORE_CUTOFF: {ADDRESS_MATCH_SCORE_CUTOFF}")
    print()
    try:
        validate()
        print("Validation passed.")
    except ValueError as e:
        print(f"Validation FAILED:\n{e}")