from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from typing import List

from services.markdown_service import parse_markdown_to_html
from services.export_service import export_document
from services.docx_transform_service import transform_docx

app = FastAPI(title="GenerateDocs")

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TEMPLATES_DIR = BASE_DIR / "templates"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <title>GenerateDocs</title>
        <link rel="stylesheet" href="/static/styles.css" />
    </head>
    <body>
        <main class="container">
            <h1>GenerateDocs</h1>
            <form action="/generate" method="post" enctype="multipart/form-data">
                <label for="title">Document title</label>
                <input id="title" name="title" type="text" placeholder="Project brief" />

                <label for="markdown">Markdown content</label>
                <textarea id="markdown" name="markdown" rows="18" placeholder="# Title\n\nWrite your content here..."></textarea>

                <label for="template">Template</label>
                <select id="template" name="template">
                    <option value="basic">Basic</option>
                    <option value="report">Report</option>
                </select>

                <label for="format">Export format</label>
                <select id="format" name="format">
                    <option value="html">HTML</option>
                    <option value="docx">DOCX</option>
                </select>

                <label for="images">Images</label>
                <input id="images" name="images" type="file" multiple accept="image/*" />

                <button type="submit">Generate document</button>
            </form>
        </main>
    </body>
    </html>
    """
    return HTMLResponse(content=html)


@app.post("/generate")
async def generate_document(
    title: str = Form(...),
    markdown: str = Form(...),
    template: str = Form("basic"),
    format: str = Form("html"),
    images: List[UploadFile] = File(default=[]),
):
    saved_images = []
    for image in images:
        if image.filename:
            file_path = UPLOAD_DIR / image.filename
            contents = await image.read()
            file_path.write_bytes(contents)
            saved_images.append(str(file_path))

    html_content = parse_markdown_to_html(markdown, title, template)
    output_file = export_document(html_content, title, format, saved_images)
    return FileResponse(path=str(output_file), filename=output_file.name, media_type="application/octet-stream")


@app.post("/transform-docx")
async def transform_uploaded_docx(
    source: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    source_path = UPLOAD_DIR / "source_input.docx"
    template_path = UPLOAD_DIR / "layout_template.docx"
    output_path = OUTPUT_DIR / "formatted_document.docx"
    source_path.write_bytes(await source.read())
    template_path.write_bytes(await template_file.read())
    transform_docx(source_path, template_path, output_path)
    return FileResponse(
        path=str(output_path),
        filename=output_path.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
