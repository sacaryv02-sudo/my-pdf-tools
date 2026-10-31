import os
import shutil
import subprocess
import uuid
from typing import List

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from pypdf import PdfWriter
from docx import Document
from docxcompose.composer import Composer
import pandas as pd

app = FastAPI(
    title="Document Converter & Merger API",
    description="Backend API for Word/PPT to PDF conversion and Word, Excel, and PDF merging."
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_TEMP_DIR = "temp_workspaces"
os.makedirs(BASE_TEMP_DIR, exist_ok=True)


# --- HELPER FUNCTIONS ---

def create_workspace() -> str:
    """Creates an isolated temporary folder for a single request."""
    workspace_path = os.path.join(BASE_TEMP_DIR, str(uuid.uuid4()))
    os.makedirs(workspace_path, exist_ok=True)
    return workspace_path


def cleanup_workspace(workspace_path: str):
    """Deletes the temporary workspace folder after the file is sent to the user."""
    if os.path.exists(workspace_path):
        shutil.rmtree(workspace_path, ignore_errors=True)


def save_upload(upload_file: UploadFile, workspace: str) -> str:
    """Saves an uploaded file to the workspace and returns its file path."""
    safe_filename = f"{uuid.uuid4().hex[:8]}_{upload_file.filename}"
    file_path = os.path.join(workspace, safe_filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)
        # --- 1. SERVE YOUR FRONTEND (index.html) AT THE MAIN URL ---
@app.get("/")
async def serve_home():
    if os.path.exists("index.html"):
        return FileResponse("index.html")
    return {"message": "Document Converter & Merger API is running! Go to /docs to test endpoints."}


# --- 2. CONVERT OFFICE FILES (WORD, EXCEL, PPT) TO PDF ---
@app.post("/convert-to-pdf")
async def convert_to_pdf(
    background_tasks: BackgroundTasks, 
    file: UploadFile = File(...)
):
    temp_dir = create_request_temp_dir()
    input_path = os.path.join(temp_dir, file.filename)

    # Save uploaded file to temporary directory
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Convert to PDF using LibreOffice headless mode
    try:
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", temp_dir, input_path],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
    except Exception as e:
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    base_name = os.path.splitext(file.filename)[0]
    output_filename = f"{base_name}.pdf"
    output_path = os.path.join(temp_dir, output_filename)

    if not os.path.exists(output_path):
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail="Output PDF was not generated.")

    return send_file_and_cleanup(output_path, output_filename, temp_dir, background_tasks)


# --- 3. MERGE MULTIPLE PDF FILES ---
@app.post("/merge-pdfs")
async def merge_pdfs(
    background_tasks: BackgroundTasks, 
    files: List[UploadFile] = File(...)
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 PDF files to merge.")

    temp_dir = create_request_temp_dir()
    merged_pdf = fitz.open()  # PyMuPDF

    try:
        for file in files:
            if not file.filename.lower().endswith(".pdf"):
                raise HTTPException(status_code=400, detail=f"File {file.filename} is not a PDF.")
            
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as buffer:
                sh
