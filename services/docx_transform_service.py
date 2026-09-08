from copy import deepcopy
from pathlib import Path

from docx import Document


def transform_docx(source_path: str | Path, template_path: str | Path, output_path: str | Path) -> Path:
    """Create a new DOCX using template layout and source document content."""
    source = Document(str(source_path))
    output = Document(str(template_path))

    body = output._element.body
    for child in list(body):
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)

    for element in source._element.body:
        if element.tag.endswith("}sectPr"):
            continue
        body.insert(len(body) - 1, deepcopy(element))

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    output.save(str(destination))
    return destination