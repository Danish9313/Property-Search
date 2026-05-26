"""
Debt Recovery Analysis — FastAPI
=================================
Upload your Excel file + all JSON files from your property folder.
The pipeline runs automatically and returns a formatted Excel report.

Start the server:
    venv/Scripts/uvicorn api:app --reload --port 8000

Interactive UI (drag & drop files here):
    http://localhost:8000/docs
"""

import os
import uuid
import shutil
import zipfile
import logging
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

import config
from main import (
    CustomerDataStore,
    process_row,
    write_results,
    AI_OUTPUT_COLUMNS,
)

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
# App
# ---------------------------------------------------------------------------
STATIC_DIR = Path(r"C:\Users\DHussain\OneDrive - Cedar Financial\Desktop\property search\static")

app = FastAPI(
    title="Debt Recovery Analysis API",
    description=(
        "**How to use:**\n\n"
        "1. Go to `POST /analyze` below and click **Try it out**\n"
        "2. Upload your **Excel file** (`.xlsx` or `.ods`)\n"
        "3. ZIP your JSON folder → upload that single **ZIP file**\n"
        "4. Click **Execute** — you get a `job_id`\n"
        "5. Check progress at `GET /jobs/{job_id}`\n"
        "6. Download your result at `GET /jobs/{job_id}/download`"
    ),
    version="1.0.0",
)

# Temp folder for job files
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

TEMP_DIR = Path(r"C:\Users\DHussain\OneDrive - Cedar Financial\Desktop\property search\temp_jobs")
TEMP_DIR.mkdir(exist_ok=True)

# In-memory job tracker
JOBS: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# Pipeline runner (runs in background)
# ---------------------------------------------------------------------------
def run_pipeline(
    job_id: str,
    excel_path: str,
    json_dir: str,
    output_path: str,
    api_key: str,
    max_rows: Optional[int],
):
    try:
        JOBS[job_id]["status"] = "running"
        JOBS[job_id]["message"] = "Loading Excel file..."

        # Load Excel
        ext = Path(excel_path).suffix.lower()
        df = pd.read_excel(excel_path, engine="odf", dtype=str) if ext == ".ods" else pd.read_excel(excel_path, dtype=str)
        df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

        if max_rows:
            df = df.head(max_rows)

        total = len(df)
        JOBS[job_id]["total_rows"] = total
        JOBS[job_id]["message"] = f"Loaded {total} rows. Starting analysis..."
        logger.info(f"[{job_id}] {total} rows loaded.")

        # Load JSON store
        json_store = CustomerDataStore(json_dir)

        # Process each row
        all_results: list[dict] = []
        for idx, row in df.iterrows():
            JOBS[job_id]["processed_rows"] = idx + 1
            JOBS[job_id]["message"] = f"Analyzing row {idx + 1} of {total}  —  CUST: {row.get('CUST_NUMBER', '?')}"
            logger.info(f"[{job_id}] Row {idx + 1}/{total} | CUST={row.get('CUST_NUMBER', '?')}")

            try:
                result = process_row(row, json_store, api_key)
                if result.get("_skip"):
                    result = {col: "" for col in AI_OUTPUT_COLUMNS}
                    result["notes"] = "Skipped — Property Status != 'property found'"
            except Exception as e:
                logger.error(f"[{job_id}] Row {idx} error: {e}")
                result = {col: "" for col in AI_OUTPUT_COLUMNS}
                result["notes"] = f"PIPELINE ERROR: {e}"

            all_results.append(result)

        # Write output Excel
        write_results(df, all_results, output_path)

        JOBS[job_id]["status"] = "done"
        JOBS[job_id]["message"] = f"Complete — {total} rows processed."
        JOBS[job_id]["output_path"] = output_path
        logger.info(f"[{job_id}] Done. Output: {output_path}")

    except Exception as e:
        logger.error(f"[{job_id}] Pipeline failed: {e}")
        JOBS[job_id]["status"] = "error"
        JOBS[job_id]["message"] = str(e)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", include_in_schema=False)
def root():
    html = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@app.get("/health", summary="Health check")
def health():
    return {
        "status": "ok",
        "model": config.GEMINI_MODEL,
        "api_key_set": bool(config.GEMINI_API_KEY),
    }


@app.post("/analyze", summary="Upload Excel + JSON ZIP → download result Excel")
async def analyze(
    background_tasks: BackgroundTasks,
    excel_file: UploadFile = File(..., description="Customer data file — .xlsx or .ods"),
    json_zip:   UploadFile = File(..., description="ZIP of your JSON folder — right-click folder → Send to → Compressed (zipped)"),
    max_rows:   Optional[int] = Form(None, description="Limit rows for testing (leave blank = process all rows)"),
):
    """
    Upload two files and get the result Excel back directly as a download.

    - `excel_file` → your customer Excel or ODS file
    - `json_zip` → ZIP your **Json_Results folder** (right-click → Send to → Compressed), upload here
    - `max_rows` → optional, limits rows for a quick test run

    The pipeline runs and **returns `debt_recovery_output.xlsx` immediately as a download.**
    """

    # Validate Excel
    excel_name = excel_file.filename or ""
    if not excel_name.lower().endswith((".xlsx", ".xls", ".ods")):
        raise HTTPException(400, "excel_file must be .xlsx, .xls, or .ods")

    # Validate ZIP
    zip_name = json_zip.filename or ""
    if not zip_name.lower().endswith(".zip"):
        raise HTTPException(400, "json_zip must be a .zip file")

    if not config.GEMINI_API_KEY:
        raise HTTPException(500, "GEMINI_API_KEY is not configured on the server.")

    # Create unique workspace
    job_id   = str(uuid.uuid4())[:8]
    job_dir  = TEMP_DIR / job_id
    json_dir = job_dir / "Json_Results"
    json_dir.mkdir(parents=True)

    # Save Excel
    excel_path = job_dir / excel_name
    excel_path.write_bytes(await excel_file.read())

    # Save and extract ZIP
    zip_path = job_dir / "uploaded.zip"
    zip_path.write_bytes(await json_zip.read())

    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.namelist():
            if member.startswith("__MACOSX") or not member.endswith(".txt"):
                continue
            dest = json_dir / member
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(zf.read(member))

    txt_count = sum(1 for _ in json_dir.rglob("*.txt"))
    if txt_count == 0:
        shutil.rmtree(job_dir)
        raise HTTPException(400, "No .txt property files found inside the ZIP. Make sure you zipped the correct folder.")

    output_path = str(job_dir / "debt_recovery_output.xlsx")

    logger.info(f"[{job_id}] Starting pipeline — {excel_name}, {txt_count} TXT property files.")

    # Run pipeline synchronously so we can return the file directly
    try:
        ext = Path(str(excel_path)).suffix.lower()
        df = pd.read_excel(str(excel_path), engine="odf", dtype=str) if ext == ".ods" else pd.read_excel(str(excel_path), dtype=str)
        df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

        if max_rows:
            df = df.head(max_rows)

        total = len(df)
        logger.info(f"[{job_id}] {total} rows loaded.")

        json_store = CustomerDataStore(str(json_dir))

        all_results: list[dict] = []
        for idx, row in df.iterrows():
            logger.info(f"[{job_id}] Row {idx + 1}/{total} | CUST={row.get('CUST_NUMBER', '?')}")
            try:
                result = process_row(row, json_store, config.GEMINI_API_KEY)
                if result.get("_skip"):
                    result = {col: "" for col in AI_OUTPUT_COLUMNS}
                    result["notes"] = "Skipped — Property Status != 'property found'"
            except Exception as e:
                logger.error(f"[{job_id}] Row {idx} error: {e}")
                result = {col: "" for col in AI_OUTPUT_COLUMNS}
                result["notes"] = f"PIPELINE ERROR: {e}"
            all_results.append(result)

        write_results(df, all_results, output_path)
        logger.info(f"[{job_id}] Done — returning file.")

    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(500, f"Pipeline error: {e}")

    # Clean up temp files after response is sent
    background_tasks.add_task(shutil.rmtree, job_dir, True)

    return FileResponse(
        path       = output_path,
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename   = "debt_recovery_output.xlsx",
    )


@app.get("/jobs/{job_id}", summary="Check job progress")
def job_status(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, f"Job '{job_id}' not found.")

    job = JOBS[job_id].copy()
    job.pop("output_path", None)

    total     = job.get("total_rows") or 0
    processed = job.get("processed_rows") or 0
    job["progress"] = f"{processed}/{total} rows" if total else "starting..."
    job["progress_pct"] = f"{round(processed / total * 100)}%" if total else "0%"

    return job


@app.get("/jobs/{job_id}/download", summary="Download result Excel (when status = done)")
def download_result(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, f"Job '{job_id}' not found.")

    job = JOBS[job_id]

    if job["status"] == "error":
        raise HTTPException(500, f"Job failed: {job['message']}")
    if job["status"] != "done":
        raise HTTPException(202, f"Not ready — {job['message']} (status: {job['status']})")

    output_path = job.get("output_path")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(500, "Output file not found.")

    return FileResponse(
        path       = output_path,
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename   = "debt_recovery_output.xlsx",
    )


@app.get("/jobs", summary="List all active jobs")
def list_jobs():
    return {
        jid: {
            "status":   j["status"],
            "message":  j["message"],
            "progress": f"{j.get('processed_rows',0)}/{j.get('total_rows') or '?'} rows",
        }
        for jid, j in JOBS.items()
    }


@app.delete("/jobs/{job_id}", summary="Delete job and clean up files")
def delete_job(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, f"Job '{job_id}' not found.")
    job_dir = TEMP_DIR / job_id
    if job_dir.exists():
        shutil.rmtree(job_dir)
    del JOBS[job_id]
    return {"message": f"Job '{job_id}' deleted."}
