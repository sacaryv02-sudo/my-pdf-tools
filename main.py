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

# Enable CORS so your frontend (Next.js / React) can communicate with this backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Replace "*" with your frontend domain in production
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
    return file_path


def convert_office_to_pdf(input_path: str, output_dir: str) -> str:
    """Converts a .docx or .pptx file to .pdf using LibreOffice Headless."""
    # On Windows, you may need to replace "soffice" with the full path to soffice.exe
    cmd = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        raise HTTPException(
            status_code=500,
            detail=f"LibreOffice conversion failed: {result.stderr.decode('utf-8', errors='ignore')}"
        )

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")

    if not os.path.exists(pdf_path):
        raise HTTPException(status_code=500, detail="Output PDF file was not generated.")
    return pdf_path


# --- API ENDPOINTS ---

@app.post("/api/word-to-pdf")
def word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """1. Converts a Word (.doc / .docx) file into a PDF."""
    if not file.filename.lower().endswith((".doc", ".docx")):
        raise HTTPException(status_code=400, detail="Please upload a valid .doc or .docx file.")

    workspace = create_workspace()
    background_tasks.add_task(cleanup_workspace, workspace)

    input_path = save_upload(file, workspace)
    pdf_path = convert_office_to_pdf(input_path, workspace)

    output_name = f"{os.path.splitext(file.filename)[0]}.pdf"
    return FileResponse(pdf_path, media_type="application/pdf", filename=output_name)


@app.post("/api/ppt-to-pdf")
def ppt_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """2. Converts a PowerPoint (.ppt / .pptx) file into a PDF."""
    if not file.filename.lower().endswith((".ppt", ".pptx")):
        raise HTTPException(status_code=400, detail="Please upload a valid .ppt or .pptx file.")

    workspace = create_workspace()
    background_tasks.add_task(cleanup_workspace, workspace)

    input_path = save_upload(file, workspace)
    pdf_path = convert_office_to_pdf(input_path, workspace)

    output_name = f"{os.path.splitext(file.filename)[0]}.pdf"
    return FileResponse(pdf_path, media_type="application/pdf", filename=output_name)


@app.post("/api/merge-word-and-pdf")
def merge_word_and_pdf(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    """3. Merges any combination of Word (.docx) and PDF (.pdf) files into one PDF."""
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Upload at least 2 files to merge.")

    workspace = create_workspace()
    background_tasks.add_task(cleanup_workspace, workspace)

    pdf_paths = []
    for file in files:
        ext = os.path.splitext(file.filename.lower())[1]
        saved_path = save_upload(file, workspace)

        if ext in [".doc", ".docx"]:
            # Convert Word file to PDF first so it can be merged cleanly
            converted_pdf = convert_office_to_pdf(saved_path, workspace)
            pdf_paths.append(converted_pdf)
        elif ext == ".pdf":
            pdf_paths.append(saved_path)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {file.filename}")

    # Merge all PDFs in the exact order uploaded
    merger = PdfWriter()
    for pdf in pdf_paths:
        merger.append(pdf)

    output_pdf_path = os.path.join(workspace, "merged_document.pdf")
    merger.write(output_pdf_path)
    merger.close()

    return FileResponse(output_pdf_path, media_type="application/pdf", filename="merged_document.pdf")


@app.post("/api/merge-word")
def merge_word_files(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    """4. Merges multiple Word (.docx) files into a single Word (.docx) file while preserving styles."""
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Upload at least 2 .docx files to merge.")

    workspace = create_workspace()
    background_tasks.add_task(cleanup_workspace, workspace)

    saved_paths = []
    for file in files:
        if not file.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="All files must be in .docx format.")
        saved_paths.append(save_upload(file, workspace))

    # Initialize the master document with the first uploaded Word file
    master_doc = Document(saved_paths[0])
    master_doc.add_page_break()
    composer = Composer(master_doc)

    # Append subsequent Word files
    for i, doc_path in enumerate(saved_paths[1:]):
        sub_doc = Document(doc_path)
        if i < len(saved_paths[1:]) - 1:
            sub_doc.add_page_break()
        composer.append(sub_doc)

    output_docx_path = os.path.join(workspace, "merged_word.docx")
    composer.save(output_docx_path)

    return FileResponse(
        output_docx_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="merged_word.docx"
    )


@app.post("/api/merge-excel")
def merge_excel_files(
    background_tasks: BackgroundTasks,
    mode: str = "rows",  # Options: "rows" (stack into 1 sheet) or "sheets" (separate tabs)
    files: List[UploadFile] = File(...)
):
    """5. Merges multiple Excel (.xlsx) files either by stacking rows or combining into separate sheets."""
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Upload at least 2 Excel files to merge.")

    workspace = create_workspace()
    background_tasks.add_task(cleanup_workspace, workspace)

    saved_paths = []
    for file in files:
        if not file.filename.lower().endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="All files must be Excel (.xlsx or .xls) files.")
        saved_paths.append((file.filename, save_upload(file, workspace)))

    output_excel_path = os.path.join(workspace, "merged_excel.xlsx")

    if mode == "sheets":
        # Mode A: Each uploaded Excel file becomes its own worksheet tab
        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            for idx, (orig_name, path) in enumerate(saved_paths):
                df = pd.read_excel(path)
                # Excel sheet names are limited to 31 characters
                sheet_name = f"{idx+1}_{os.path.splitext(orig_name)[0]}"[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Mode B (Default): Stack all rows into one master sheet
        dataframes = [pd.read_excel(path) for _, path in saved_paths]
        merged_df = pd.concat(dataframes, ignore_index=True)
        merged_df.to_excel(output_excel_path, index=False, engine="openpyxl")

    return FileResponse(
        output_excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="merged_excel.xlsx"
    )
