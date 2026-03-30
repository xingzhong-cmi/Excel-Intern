"""FastAPI application for Excel Assistant."""

import logging
import os
import uuid
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from backend.script_generator import generate_and_execute

# Load environment variables
load_dotenv()
load_dotenv("config/.env")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# Directories
UPLOADS_DIR = Path("uploads")
RESULTS_DIR = Path("results")
TEMP_DIR = Path("temp")

for d in [UPLOADS_DIR, RESULTS_DIR, TEMP_DIR]:
    d.mkdir(exist_ok=True)

# FastAPI app
app = FastAPI(title="Excel助手", version="1.0.0")

# CORS - configurable via CORS_ORIGINS env var (comma-separated), defaults to * for development
cors_origins = os.getenv("CORS_ORIGINS", "*").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve frontend static files
FRONTEND_DIR = Path("frontend")
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory="frontend"), name="static")


# --- Models ---


class ProcessRequest(BaseModel):
    """Request model for processing instruction."""

    filenames: list[str]
    instruction: str


# --- API Endpoints ---


@app.get("/")
async def root():
    """Serve the main frontend page."""
    index_path = FRONTEND_DIR / "index.html"
    if index_path.exists():
        return FileResponse(str(index_path), media_type="text/html")
    return {"message": "Excel助手 API is running. Frontend not found."}


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    Upload an Excel file and return a preview.

    Accepts .xlsx, .xls, and .csv files.
    """
    # Validate file extension
    allowed_extensions = {".xlsx", ".xls", ".csv"}
    suffix = Path(file.filename).suffix.lower()
    if suffix not in allowed_extensions:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file type: {suffix}. Allowed: {', '.join(allowed_extensions)}",
        )

    # Save with unique prefix to avoid collisions
    safe_filename = f"{uuid.uuid4().hex[:8]}_{file.filename}"
    filepath = UPLOADS_DIR / safe_filename

    try:
        content = await file.read()
        filepath.write_bytes(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {e}")

    # Read and return preview
    try:
        if suffix == ".csv":
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)

        preview_data = df.head(50).fillna("").to_dict(orient="records")

        return {
            "filename": safe_filename,
            "original_name": file.filename,
            "rows": len(df),
            "columns": list(df.columns),
            "preview": preview_data,
        }
    except Exception as e:
        filepath.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=f"Failed to read Excel file: {e}")


@app.post("/api/process")
async def process_instruction(request: ProcessRequest):
    """
    Process a natural language instruction on the uploaded files.

    Generates a Python script via LLM, validates it, executes it,
    and returns the result preview.
    """
    if not request.filenames:
        raise HTTPException(status_code=400, detail="No files specified.")
    if not request.instruction.strip():
        raise HTTPException(status_code=400, detail="Instruction cannot be empty.")

    # Verify files exist
    for filename in request.filenames:
        if not (UPLOADS_DIR / filename).exists():
            raise HTTPException(
                status_code=404, detail=f"File not found: {filename}"
            )

    result = await generate_and_execute(request.filenames, request.instruction)
    return result


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Download a processed result file."""
    filepath = RESULTS_DIR / filename
    if not filepath.exists():
        raise HTTPException(status_code=404, detail="File not found.")

    # Prevent path traversal
    if not filepath.resolve().is_relative_to(RESULTS_DIR.resolve()):
        raise HTTPException(status_code=403, detail="Access denied.")

    return FileResponse(
        str(filepath),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/preview/{filename}")
async def preview_file(filename: str):
    """Get a preview of an uploaded or result file."""
    # Check in uploads first, then results
    filepath = UPLOADS_DIR / filename
    if not filepath.exists():
        filepath = RESULTS_DIR / filename
    if not filepath.exists():
        raise HTTPException(status_code=404, detail="File not found.")

    # Prevent path traversal
    if not (
        filepath.resolve().is_relative_to(UPLOADS_DIR.resolve())
        or filepath.resolve().is_relative_to(RESULTS_DIR.resolve())
    ):
        raise HTTPException(status_code=403, detail="Access denied.")

    try:
        suffix = filepath.suffix.lower()
        if suffix == ".csv":
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath)

        return {
            "filename": filename,
            "rows": len(df),
            "columns": list(df.columns),
            "preview": df.head(50).fillna("").to_dict(orient="records"),
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to read file: {e}")
