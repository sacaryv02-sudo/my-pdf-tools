import os
import uuid
import shutil
import subprocess
from typing import List

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.openapi.utils import get_openapi

from pypdf import PdfReader, PdfWriter
from docx import Document
from docxcompose.composer import Composer
import openpyxl
from pdf2docx import Converter

app = FastAPI(
    title="Document Converter & Merger API",
    description="Backend API for PDF/PPTX/DOCX/XLSX conversion and Word, Excel, and PDF merging.",
    version="1.0.0"
)

# Enable CORS so index.html can communicate with the API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "/tmp/doc_processing"
os.makedirs(TEMP_DIR, exist_ok=True)


# --- HELPER FUNCTIONS ---

def cleanup_files(file_paths: List[str]):
    """Deletes temporary files and folders after the response is sent."""
    for path in file_paths:
        try:
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            elif os.path.exists(path):
                os.remove(path)
        except Exception as e:
            print(f"Cleanup error for {path}: {e}")


def save_upload_file(upload_file: UploadFile, destination: str):
    """Saves an uploaded file to disk."""
    with open(destination, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)


# --- SERVE FRONTEND WEBSITE ---

@app.get("/", include_in_schema=False)
async def serve_homepage():
    if os.path.exists("index.html"):
        return FileResponse("index.html")
    return {"message": "Document Converter & Merger API is running. Visit /docs for the API UI."}


# --- 1. CONVERT PDF TO WORD (.DOCX) ---

@app.post("/convert-pdf-to-word")
async def convert_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="Select a PDF file to convert to Word")
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a PDF (.pdf)")

    job_id = str(uuid.uuid4())
    input_path = os.path.join(TEMP_DIR, f"{job_id}_input.pdf")
    output_filename = f"{os.path.splitext(file.filename)[0]}_converted.docx"
    output_path = os.path.join(TEMP_DIR, f"{job_id}_{output_filename}")

    try:
        save_upload_file(file, input_path)
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="PDF to Word conversion failed.")

        background_tasks.add_task(cleanup_files, [input_path, output_path])
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        cleanup_files([input_path, output_path])
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# --- 2. CONVERT OFFICE FILES (WORD / EXCEL / PPT) TO PDF ---

@app.post("/convert-to-pdf")
async def convert_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="Select a Word, Excel, or PowerPoint file")
):
    allowed_exts = (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt")
    if not file.filename.lower().endswith(allowed_exts):
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format. Allowed formats: {', '.join(allowed_exts)}"
        )

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    input_path = os.path.join(job_dir, file.filename)
    output_filename = f"{os.path.splitext(file.filename)[0]}.pdf"
    output_path = os.path.join(job_dir, output_filename)

    try:
        save_upload_file(file, input_path)
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", job_dir,
            input_path
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if result.returncode != 0 or not os.path.exists(output_path):
            raise HTTPException(
                status_code=500,
                detail=f"LibreOffice conversion failed: {result.stderr}"
            )

        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/pdf"
        )
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# --- 3. MERGE MULTIPLE PDF FILES ---

@app.post("/merge-pdfs")
async def merge_pdfs(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(..., description="Select 2 or more PDF files to merge")
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 PDF files to merge.")

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    output_path = os.path.join(job_dir, "merged_document.pdf")

    writer = PdfWriter()

    try:
        for idx, file in enumerate(files):
            if not file.filename.lower().endswith(".pdf"):
                raise HTTPException(
                    status_code=400,
                    detail=f"File '{file.filename}' is not a PDF."
                )
            file_path = os.path.join(job_dir, f"{idx}_{file.filename}")
            save_upload_file(file, file_path)

            reader = PdfReader(file_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as out_file:
            writer.write(out_file)

        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename="merged_document.pdf",
            media_type="application/pdf"
        )
    except HTTPException:
        cleanup_files([job_dir])
        raise
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"PDF Merge error: {str(e)}")


# --- 4. MERGE MULTIPLE WORD (.DOCX) FILES ---

@app.post("/merge-word")
async def merge_word(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(..., description="Select 2 or more Word (.docx) files to merge")
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 Word (.docx) files to merge.")

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    output_path = os.path.join(job_dir, "merged_word.docx")

    try:
        saved_paths = []
        for idx, file in enumerate(files):
            if not file.filename.lower().endswith(".docx"):
                raise HTTPException(
                    status_code=400,
                    detail=f"File '{file.filename}' must be a .docx file."
                )
            file_path = os.path.join(job_dir, f"{idx}_{file.filename}")
            save_upload_file(file, file_path)
            saved_paths.append(file_path)

        master_doc = Document(saved_paths[0])
        composer = Composer(master_doc)

        for path in saved_paths[1:]:
            master_doc.add_page_break()
            sub_doc = Document(path)
            composer.append(sub_doc)

        composer.save(output_path)

        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename="merged_word.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except HTTPException:
        cleanup_files([job_dir])
        raise
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Word Merge error: {str(e)}")


# --- 5. MERGE MULTIPLE EXCEL (.
