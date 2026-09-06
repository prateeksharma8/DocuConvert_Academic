# GenerateDocs

A minimal document generation app that turns markdown and images into polished downloadable outputs.

## Features

- Markdown input via textarea or file upload
- Image upload support
- Template selection
- Export to HTML or DOCX
- Download generated file

## Tech stack

- FastAPI
- Jinja2
- Python-Markdown
- python-docx

## Local setup

1. Create a virtual environment
2. Install dependencies
3. Run the app

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Then open http://localhost:8000

## Project structure

- app/
- services/
- templates/
- static/
- uploads/
- visualizations/ (generated assets only)

## CLI Validation

You can validate a DOCX file locally using the provided `validate_doc.py` script.

Example:

```bash
python validate_doc.py --file path/to/paper.docx --template context.yaml --out report.json
```

The script prints a short summary and can write a detailed JSON report.
