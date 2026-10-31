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
