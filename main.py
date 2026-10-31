import os
import shutil
import uuid
import subprocess
from typing import List

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse

from pypdf import PdfWriter, PdfReader
from docx import Document
from docxcompose.composer import Composer
import openpyxl
import pandas as pd
from pdf2docx import Converter

app = FastAPI(
    title="Document Converter & Merger API",
    description="Backend API for PDF/PPTX/DOCX/XLSX conversion and Word, Excel, and PDF merging.",
    version="1.0.0"
)

TEMP_BASE_DIR = "/tmp/doc_processing"
os.makedirs(TEMP_BASE_DIR, exist_ok=True)


# --- HELPER FUNCTIONS ---
def create_request_temp_dir() -> str:
    req_dir = os.path.join(TEMP_BASE_DIR, str(uuid.uuid4()))
    os.makedirs(req_dir, exist_ok=True)
    return req_dir


def cleanup_temp_dir(dir_path: str):
    if os.path.exists(dir_path):
        shutil.rmtree(dir_path, ignore_errors=True)


def send_file_and_cleanup(file_path: str, filename: str, temp_dir: str, background_tasks: BackgroundTasks):
    background_tasks.add_task(cleanup_temp_dir, temp_dir)
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/octet-stream"
    )


# --- 1. CONVERT PDF TO WORD (DOCX) ---
@app.post("/convert-pdf-to-word")
async def convert_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a PDF.")

    temp_dir = create_request_temp_dir()
    input_path = os.path.join(temp_dir, file.filename)

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    base_name = os.path.splitext(file.filename)[0]
    output_filename = f"{base_name}.docx"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
    except Exception as e:
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail=f"PDF to Word conversion failed: {str(e)}")

    if not os.path.exists(output_path):
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail="Output Word file was not generated.")

    return send_file_and_cleanup(output_path, output_filename, temp_dir, background_tasks)


# --- 2. CONVERT OFFICE FILES (WORD, EXCEL, PPT) TO PDF ---
@app.post("/convert-to-pdf")
async def convert_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    temp_dir = create_request_temp_dir()
    input_path = os.path.join(temp_dir, file.filename)

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

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
    merger = PdfWriter()

    try:
        for idx, file in enumerate(files):
            if not file.filename.lower().endswith(".pdf"):
                raise HTTPException(status_code=400, detail=f"File {file.filename} is not a PDF.")

            file_path = os.path.join(temp_dir, f"{idx}_{file.filename}")
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            reader = PdfReader(file_path)
            for page in reader.pages:
                merger.add_page(page)

        output_filename = "merged_document.pdf"
        output_path = os.path.join(temp_dir, output_filename)
        with open(output_path, "wb") as out_file:
            merger.write(out_file)
        merger.close()

    except HTTPException:
        cleanup_temp_dir(temp_dir)
        raise
    except Exception as e:
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail=f"PDF merge failed: {str(e)}")

    return send_file_and_cleanup(output_path, output_filename, temp_dir, background_tasks)


# --- 4. MERGE MULTIPLE WORD (DOCX) FILES ---
@app.post("/merge-word")
async def merge_word(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...)
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 Word (.docx) files to merge.")

    temp_dir = create_request_temp_dir()
    saved_paths = []

    try:
        for idx, file in enumerate(files):
            if not file.filename.lower().endswith(".docx"):
                raise HTTPException(status_code=400, detail=f"File {file.filename} is not a .docx file.")

            file_path = os.path.join(temp_dir, f"{idx}_{file.filename}")
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            saved_paths.append(file_path)

        master_doc = Document(saved_paths[0])
        composer = Composer(master_doc)

        for path in saved_paths[1:]:
            doc_to_append = Document(path)
            composer.append(doc_to_append)

        output_filename = "merged_document.docx"
        output_path = os.path.join(temp_dir, output_filename)
        composer.save(output_path)

    except HTTPException:
        cleanup_temp_dir(temp_dir)
        raise
    except Exception as e:
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail=f"Word merge failed: {str(e)}")

    return send_file_and_cleanup(output_path, output_filename, temp_dir, background_tasks)


# --- 5. MERGE MULTIPLE EXCEL (XLSX) FILES ---
@app.post("/merge-excel")
async def merge_excel(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...)
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 Excel (.xlsx) files to merge.")

    temp_dir = create_request_temp_dir()
    output_filename = "merged_workbook.xlsx"
    output_path = os.path.join(temp_dir, output_filename)

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for idx, file in enumerate(files):
                if not file.filename.lower().endswith((".xlsx", ".xls")):
                    raise HTTPException(status_code=400, detail=f"File {file.filename} is not an Excel file.")

                file_path = os.path.join(temp_dir, f"{idx}_{file.filename}")
                with open(file_path, "wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)

                excel_data = pd.read_excel(file_path, sheet_name=None)
                base_name = os.path.splitext(file.filename)[0][:15]

                for sheet_name, df in excel_data.items():
                    safe_sheet_name = f"{idx+1}_{base_name}_{sheet_name}"[:31]
                    df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

    except HTTPException:
        cleanup_temp_dir(temp_dir)
        raise
    except Exception as e:
        cleanup_temp_dir(temp_dir)
        raise HTTPException(status_code=500, detail=f"Excel merge failed: {str(e)}")

    return send_file_and_cleanup(output_path, output_filename, temp_dir, background_tasks)
