import os
import time
import uuid
import shutil
import subprocess
from typing import List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
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
    upload_file.file.seek(0)
    with open(destination, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)


def convert_office_to_pdf_helper(input_path: str, output_dir: str) -> str:
    """Converts a Word, Excel, or PowerPoint file to PDF using LibreOffice."""
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    output_filename = f"{os.path.splitext(os.path.basename(input_path))[0]}.pdf"
    output_path = os.path.join(output_dir, output_filename)

    if result.returncode != 0 or not os.path.exists(output_path):
        raise Exception(f"LibreOffice conversion failed: {result.stderr}")
    return output_path


# --- SERVE FRONTEND WEBSITE ---

@app.get("/", include_in_schema=False)
async def serve_homepage():
    if os.path.exists("index.html"):
        return FileResponse("index.html")
    return {"message": "Document Converter & Merger API is running. Visit /docs for the API UI."}


# --- 1. CONVERT PDF TO WORD (.DOCX) ---

@app.post("/convert-pdf-to-word")
@app.post("/api/pdf-to-word", include_in_schema=False)
async def convert_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="Select a PDF file to convert to Word")
):
    start_time = time.time()
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

        elapsed = round(time.time() - start_time, 2)
        background_tasks.add_task(cleanup_files, [input_path, output_path])
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"X-Processing-Time-Seconds": str(elapsed)}
        )
    except Exception as e:
        cleanup_files([input_path, output_path])
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# --- 2. CONVERT WORD / PPT / EXCEL TO PDF ---

@app.post("/convert-to-pdf")
@app.post("/api/word-to-pdf", include_in_schema=False)
@app.post("/api/ppt-to-pdf", include_in_schema=False)
@app.post("/api/convert-to-pdf", include_in_schema=False)
async def convert_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="Select a Word, Excel, or PowerPoint file")
):
    start_time = time.time()
    allowed_exts = (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt")
    if not file.filename.lower().endswith(allowed_exts):
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format. Allowed: {', '.join(allowed_exts)}"
        )

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)

    input_path = os.path.join(job_dir, file.filename)
    output_filename = f"{os.path.splitext(file.filename)[0]}.pdf"

    try:
        save_upload_file(file, input_path)
        output_path = convert_office_to_pdf_helper(input_path, job_dir)

        elapsed = round(time.time() - start_time, 2)
        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/pdf",
            headers={"X-Processing-Time-Seconds": str(elapsed)}
        )
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# --- 3. MERGE WORD (.DOCX) & PDF FILES INTO ONE PDF ---

@app.post("/merge-pdfs")
@app.post("/api/merge-word-and-pdf", include_in_schema=False)
@app.post("/api/merge-pdfs", include_in_schema=False)
async def merge_word_and_pdf(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(..., description="Select 2 or more PDF or Word (.docx) files to merge into one PDF")
):
    start_time = time.time()
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 files to merge.")

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    output_path = os.path.join(job_dir, "merged_document.pdf")

    writer = PdfWriter()

    try:
        for idx, file in enumerate(files):
            ext = os.path.splitext(file.filename.lower())[1]
            raw_path = os.path.join(job_dir, f"{idx}_{file.filename}")
            save_upload_file(file, raw_path)

            # If it's a Word file, automatically convert it to PDF first!
            if ext in (".docx", ".doc"):
                pdf_path = convert_office_to_pdf_helper(raw_path, job_dir)
            elif ext == ".pdf":
                pdf_path = raw_path
            else:
                raise HTTPException(
                    status_code=400,
                    detail=f"Unsupported file '{file.filename}'. Only .pdf, .docx, and .doc are allowed."
                )

            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as out_file:
            writer.write(out_file)

        elapsed = round(time.time() - start_time, 2)
        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename="merged_document.pdf",
            media_type="application/pdf",
            headers={"X-Processing-Time-Seconds": str(elapsed)}
        )
    except HTTPException:
        cleanup_files([job_dir])
        raise
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Merge error: {str(e)}")


# --- 4. MERGE MULTIPLE WORD (.DOCX) FILES ---

@app.post("/merge-word")
@app.post("/api/merge-word", include_in_schema=False)
async def merge_word(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(..., description="Select 2 or more Word (.docx) files to merge")
):
    start_time = time.time()
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

        elapsed = round(time.time() - start_time, 2)
        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename="merged_word.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"X-Processing-Time-Seconds": str(elapsed)}
        )
    except HTTPException:
        cleanup_files([job_dir])
        raise
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Word Merge error: {str(e)}")


# --- 5. MERGE MULTIPLE EXCEL (.XLSX) FILES ---

@app.post("/merge-excel")
@app.post("/api/merge-excel", include_in_schema=False)
async def merge_excel(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(..., description="Select 2 or more Excel (.xlsx) files to merge"),
    mode: str = Form("tabs")
):
    start_time = time.time()
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Please upload at least 2 Excel (.xlsx) files to merge.")

    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    output_path = os.path.join(job_dir, "merged_excel.xlsx")

    try:
        merged_wb = openpyxl.Workbook()

        if mode == "rows":
            # Stack all rows into a single master sheet
            master_sheet = merged_wb.active
            master_sheet.title = "Merged_Data"
            for idx, file in enumerate(files):
                if not file.filename.lower().endswith((".xlsx", ".xls")):
                    raise HTTPException(status_code=400, detail=f"File '{file.filename}' must be an Excel file.")
                file_path = os.path.join(job_dir, f"{idx}_{file.filename}")
                save_upload_file(file, file_path)

                wb = openpyxl.load_workbook(file_path, data_only=True)
                for sheet_name in wb.sheetnames:
                    source_sheet = wb[sheet_name]
                    for row_idx, row in enumerate(source_sheet.iter_rows(values_only=True)):
                        # Skip header row on subsequent files if stacking
                        if idx > 0 and row_idx == 0:
                            continue
                        master_sheet.append(list(row))
        else:
            # Separate tabs mode
            merged_wb.remove(merged_wb.active)
            for idx, file in enumerate(files):
                if not file.filename.lower().endswith((".xlsx", ".xls")):
                    raise HTTPException(status_code=400, detail=f"File '{file.filename}' must be an Excel file.")
                file_path = os.path.join(job_dir, f"{idx}_{file.filename}")
                save_upload_file(file, file_path)

                wb = openpyxl.load_workbook(file_path, data_only=True)
                base_name = os.path.splitext(file.filename)[0][:15]

                for sheet_name in wb.sheetnames:
                    source_sheet = wb[sheet_name]
                    new_title = f"{idx+1}_{base_name}_{sheet_name}"[:31]
                    target_sheet = merged_wb.create_sheet(title=new_title)

                    for row in source_sheet.iter_rows(values_only=True):
                        target_sheet.append(list(row))

        merged_wb.save(output_path)

        elapsed = round(time.time() - start_time, 2)
        background_tasks.add_task(cleanup_files, [job_dir])
        return FileResponse(
            path=output_path,
            filename="merged_excel.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"X-Processing-Time-Seconds": str(elapsed)}
        )
    except HTTPException:
        cleanup_files([job_dir])
        raise
    except Exception as e:
        cleanup_files([job_dir])
        raise HTTPException(status_code=500, detail=f"Excel Merge error: {str(e)}")


# --- FIX SWAGGER UI MULTI-FILE UPLOAD RENDERING ---

def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema
    openapi_schema = get_openapi(
        title=app.title,
        version=app.version,
        description=app.description,
        routes=app.routes,
    )
    openapi_schema["openapi"] = "3.0.2"
    for schema in openapi_schema.get("components", {}).get("schemas", {}).values():
        props = schema.get("properties", {})
        if "files" in props and props["files"].get("type") == "array":
            props["files"]["items"] = {"type": "string", "format": "binary"}
    app.openapi_schema = openapi_schema
    return app.openapi_schema

app.openapi = custom_openapi
