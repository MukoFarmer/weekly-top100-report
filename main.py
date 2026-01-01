from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import os
import uuid

from analysis import analyze
from report_writer import write_weekly_report

app = FastAPI(title="Weekly Top100 Report Generator")

# HTML templates
templates = Jinja2Templates(directory="templates")

# Output folder
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )


@app.post("/analyze")
async def analyze_files(
    raw_file: UploadFile = File(...),
    progress_file: UploadFile = File(...),
    gms_file: UploadFile = File(...)
):
    # Unique working directory per request
    request_id = str(uuid.uuid4())
    work_dir = os.path.join(OUTPUT_DIR, request_id)
    os.makedirs(work_dir, exist_ok=True)

    raw_path = os.path.join(work_dir, raw_file.filename)
    progress_path = os.path.join(work_dir, progress_file.filename)
    gms_path = os.path.join(work_dir, gms_file.filename)

    for f, p in [
        (raw_file, raw_path),
        (progress_file, progress_path),
        (gms_file, gms_path),
    ]:
        with open(p, "wb") as buffer:
            buffer.write(await f.read())

    # Run analysis
    result = analyze(
        raw_path,
        progress_path,
        gms_path,
        progress_file.filename
    )

    # Write report
    output_path = os.path.join(
        work_dir,
        f"Weekly_Top100_Report_Week_{result['week']}.docx"
    )

    write_weekly_report(result, output_path)

    return FileResponse(
        output_path,
        filename=os.path.basename(output_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
