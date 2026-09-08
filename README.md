# GenerateDocs

GenerateDocs converts Markdown or an existing DOCX into a formatted document using a DOCX layout template, then validates the result against IEEE-style rules. It reports clear pass/fail status for structure, pages, fonts, images, tables, headers, and footers.

## Elevator pitch

GenerateDocs helps researchers turn working paper content into submission-ready DOCX files and quickly identify formatting problems before sharing or submitting them.

## Features

- Markdown input via textarea or file upload
- Image upload support
- Template selection
- Export to HTML or DOCX
- Transform an existing DOCX with a layout template
- Validate one or more DOCX files with detailed reports
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

### Sample status output

The CLI prints all 12 checks with a score, status icons, and failure reasons.

Passing checks use `✅ PASS`; failed checks use `❌ FAIL`.

### Compare multiple DOCX files

Pass one template file and two or more DOCX files to compare all validation checks:

```bash
python3 compare_docs.py \
	--template context.yaml \
	--file first.docx second.docx third.docx \
	--out reports/comparison.json
```

The comparison output includes each document's score, all 12 check results side by side, document metrics, and failure reasons. The input DOCX files are read-only and are not changed.

### Create a DOCX from a source DOCX and layout template

```bash
curl -X POST http://localhost:8000/transform-docx \
	-F "source=@source.docx" \
	-F "template_file=@ieee_template.docx" \
	-o outputs/formatted_document.docx
```

The source document and layout template are unchanged. GenerateDocs creates a new output file.
