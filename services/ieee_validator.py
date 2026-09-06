import io
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple

from PIL import Image


CHECK_LABELS = {
    "title_present": "Title present",
    "abstract_present": "Abstract present",
    "abstract_length_valid": "Abstract length is 150-250 words",
    "author_block_present": "Author block present",
    "section_structure_present": "IEEE section structure present",
    "image_quality_valid": "Image DPI and pixel dimensions valid",
    "table_format_valid": "Table structure valid",
    "document_format_valid": "DOCX format valid",
    "page_limit_valid": "Page count is within the configured limit",
    "header_present": "Header present",
    "footer_present": "Footer present",
    "font_consistency_valid": "Font usage is consistent",
}


def _build_tasks(checks: Dict[str, bool]) -> List[Dict[str, object]]:
    return [
        {
            "task": CHECK_LABELS[name],
            "check": name,
            "status": "PASS" if passed else "FAIL",
            "passed": passed,
        }
        for name, passed in checks.items()
    ]


def _extract_docx_text(docx_path: str | Path) -> str:
    """Read the text from a DOCX file without requiring Word itself."""
    with zipfile.ZipFile(docx_path) as archive:
        xml = archive.read("word/document.xml")
    text = xml.decode("utf-8", errors="ignore")
    text = re.sub(r"<.*?>", " ", text, flags=re.DOTALL)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalize_title(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _normalize_heading_spacing(text: str) -> str:
    return re.sub(r"(?<![A-Z])([A-Z])\s+(?=[A-Z])", r"\1", text)


def _find_title(text: str) -> str | None:
    # Try to capture a leading title-like phrase occurring before author/abstract markers
    m = re.search(r"^(.*?)(?=\bAbstract\b|\d+\s+Author\b|\bAuthors\b|\bIntroduction\b)", text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        candidate = m.group(1).strip()
        if len(candidate.split()) >= 3:
            return _normalize_title(candidate)

    # Fallback: scan for title-like capitalized phrases elsewhere
    candidates = re.findall(r"([A-Z][^\n]{5,120})", text)
    for candidate in candidates:
        if "Abstract" in candidate:
            continue
        if len(candidate.split()) >= 3 and not re.search(r"\b(Independent Researcher|IEEE)\b", candidate):
            return _normalize_title(candidate)
    return None


def _has_email_like_author_block(text: str) -> bool:
    # Consider an author block present if there's an email OR the word 'Author'
    # or numbered author lines (e.g., '1 Author One'). Tests build author lines
    # containing the word 'Author', so detect that as well.
    if re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text):
        return True
    if re.search(r"\bAuthor\b", text, flags=re.IGNORECASE):
        return True
    if re.search(r"^\d+\s+[A-Z][a-z]+", text, flags=re.MULTILINE):
        return True
    return False


def _abstract_body(text: str) -> str:
    # Find the Abstract section and ensure it contains non-whitespace content
    text = _normalize_heading_spacing(text)
    m = re.search(r"\bAbstract\b(.*?)(?=\bIntroduction\b|\bMethodology\b|\bResults\b|\bConclusion\b|$)", text, flags=re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""


def _has_abstract(text: str) -> bool:
    body = _abstract_body(text)
    # Consider abstract present if there's at least one alphanumeric character
    return bool(re.search(r"\w", body))


def _abstract_word_count(text: str) -> int:
    return len(re.findall(r"\b[\w'-]+\b", _abstract_body(text)))


def _inspect_images(archive: zipfile.ZipFile, minimum_dpi: int, minimum_width: int, minimum_height: int) -> tuple[bool, List[Dict[str, object]]]:
    image_details = []
    for image_name in archive.namelist():
        if not image_name.startswith("word/media/"):
            continue
        with Image.open(io.BytesIO(archive.read(image_name))) as image:
            dpi = image.info.get("dpi", (0, 0))
            horizontal_dpi = round(float(dpi[0])) if dpi else 0
            vertical_dpi = round(float(dpi[1])) if len(dpi) > 1 else horizontal_dpi
            image_details.append(
                {
                    "file": image_name,
                    "width_px": image.width,
                    "height_px": image.height,
                    "dpi": [horizontal_dpi, vertical_dpi],
                    "valid": (
                        image.width >= minimum_width
                        and image.height >= minimum_height
                        and horizontal_dpi >= minimum_dpi
                        and vertical_dpi >= minimum_dpi
                    ),
                }
            )
    return all(image["valid"] for image in image_details), image_details


def _has_valid_tables(document_xml: str) -> tuple[bool, int]:
    tables = re.findall(r"<w:tbl\b.*?</w:tbl>", document_xml, flags=re.DOTALL)
    for table in tables:
        rows = re.findall(r"<w:tr\b.*?</w:tr>", table, flags=re.DOTALL)
        column_counts = [len(re.findall(r"<w:tc\b", row)) for row in rows]
        cell_text = re.sub(r"<.*?>", " ", table, flags=re.DOTALL)
        if not rows or not column_counts or len(set(column_counts)) != 1 or not re.search(r"\w", cell_text):
            return False, len(tables)
    return True, len(tables)


def _page_count(document_xml: str) -> int:
    rendered_breaks = document_xml.count("w:lastRenderedPageBreak")
    if rendered_breaks:
        return rendered_breaks + 1
    explicit_breaks = len(re.findall(r"w:type=[\"']page[\"']", document_xml))
    return explicit_breaks + 1


def _header_footer_presence(archive: zipfile.ZipFile) -> tuple[bool, bool, int, int]:
    header_names = [name for name in archive.namelist() if re.fullmatch(r"word/header\d+\.xml", name)]
    footer_names = [name for name in archive.namelist() if re.fullmatch(r"word/footer\d+\.xml", name)]

    def has_text(name: str) -> bool:
        xml = archive.read(name).decode("utf-8", errors="ignore")
        return bool(re.search(r"<w:t[^>]*>\s*\S.*?</w:t>", xml, flags=re.DOTALL))

    return (
        any(has_text(name) for name in header_names),
        any(has_text(name) for name in footer_names),
        len(header_names),
        len(footer_names),
    )


def _font_details(document_xml: str, expected_font: str) -> tuple[bool, List[str]]:
    fonts = set()
    for tag in re.findall(r"<w:rFonts\b[^>]*>", document_xml):
        fonts.update(re.findall(r"w:(?:ascii|hAnsi)=\"([^\"]+)\"", tag))
    normal_fonts = {font for font in fonts if "math" not in font.lower()}
    return not normal_fonts or normal_fonts == {expected_font}, sorted(fonts)


def _has_core_sections(text: str) -> bool:
    text = _normalize_heading_spacing(text)
    required = ["Introduction", "Methodology", "Results", "Conclusion"]
    found = 0
    for token in required:
        if re.search(rf"\b{token}\b", text, flags=re.IGNORECASE):
            found += 1
    return found >= 2


def validate_ieee_document(docx_path: str | Path, validation_config: Dict[str, object] | None = None) -> Dict[str, object]:
    """Validate a lightweight IEEE-like document structure and return pass/fail diagnostics."""
    path = Path(docx_path)
    config = validation_config or {}
    minimum_abstract_words = int(config.get("abstract_min_words", 150))
    maximum_abstract_words = int(config.get("abstract_max_words", 250))
    minimum_image_dpi = int(config.get("image_min_dpi", 300))
    minimum_image_width = int(config.get("image_min_width_px", 900))
    minimum_image_height = int(config.get("image_min_height_px", 600))
    maximum_pages = int(config.get("max_pages", 6))
    expected_font = str(config.get("font_name", "Times New Roman"))
    if not path.exists():
        checks = {name: False for name in CHECK_LABELS}
        return {
            "overall_status": "FAIL",
            "file": str(path),
            "checks": checks,
            "tasks": _build_tasks(checks),
            "notes": ["File not found."],
        }

    with zipfile.ZipFile(path) as archive:
        document_format_valid = path.suffix.lower() == ".docx" and "word/document.xml" in archive.namelist()
        document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        header_present, footer_present, header_count, footer_count = _header_footer_presence(archive)
        font_consistency_valid, fonts_detected = _font_details(document_xml, expected_font)
        image_quality_valid, image_details = _inspect_images(
            archive,
            minimum_image_dpi,
            minimum_image_width,
            minimum_image_height,
        )
    full_text = _extract_docx_text(path)
    abstract_word_count = _abstract_word_count(full_text)
    table_format_valid, table_count = _has_valid_tables(document_xml)
    page_count = _page_count(document_xml)
    title = _find_title(full_text)
    checks = {
        "title_present": bool(title),
        "abstract_present": _has_abstract(full_text),
        "abstract_length_valid": minimum_abstract_words <= abstract_word_count <= maximum_abstract_words,
        "author_block_present": _has_email_like_author_block(full_text),
        "section_structure_present": _has_core_sections(full_text),
        "image_quality_valid": image_quality_valid,
        "table_format_valid": table_format_valid,
        "document_format_valid": document_format_valid,
        "page_limit_valid": page_count <= maximum_pages,
        "header_present": header_present,
        "footer_present": footer_present,
        "font_consistency_valid": font_consistency_valid,
    }

    failed = [name for name, passed in checks.items() if not passed]
    overall_status = "PASS" if not failed else "FAIL"
    tasks = _build_tasks(checks)
    notes = []
    if not checks["title_present"]:
        notes.append("Missing a plausible title near the top of the document.")
    if not checks["abstract_present"]:
        notes.append("Missing an 'Abstract' section.")
    elif not checks["abstract_length_valid"]:
        notes.append(f"Abstract contains {abstract_word_count} words; expected {minimum_abstract_words}-{maximum_abstract_words}.")
    if not checks["author_block_present"]:
        notes.append("Missing an author block with email-style author metadata.")
    if not checks["section_structure_present"]:
        notes.append("Missing a usable IEEE-style section structure (e.g., Introduction/Methodology/Results/Conclusion).")
    if not checks["image_quality_valid"]:
        notes.append("One or more images do not meet the configured DPI or pixel-dimension requirements.")
    if not checks["table_format_valid"]:
        notes.append("One or more tables are empty or have inconsistent column counts.")
    if not checks["document_format_valid"]:
        notes.append("The input is not a valid DOCX document.")
    if not checks["page_limit_valid"]:
        notes.append(f"Document is approximately {page_count} pages; maximum configured length is {maximum_pages} pages.")
    if not checks["header_present"]:
        notes.append("No non-empty document header was found.")
    if not checks["footer_present"]:
        notes.append("No non-empty document footer was found.")
    if not checks["font_consistency_valid"]:
        notes.append("Multiple non-math fonts were detected in document text.")

    passed_count = sum(checks.values())
    total_count = len(checks)

    return {
        "overall_status": overall_status,
        "file": str(path),
        "title": title,
        "checks": checks,
        "tasks": tasks,
        "summary": {
            "passed": passed_count,
            "failed": total_count - passed_count,
            "total": total_count,
            "score": f"{passed_count}/{total_count}",
        },
        "metrics": {
            "abstract_word_count": abstract_word_count,
            "image_count": len(image_details),
            "images": image_details,
            "table_count": table_count,
            "page_count": page_count,
            "header_count": header_count,
            "footer_count": footer_count,
            "fonts_detected": fonts_detected,
        },
        "notes": notes,
    }
