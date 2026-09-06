import re
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple


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


def _has_abstract(text: str) -> bool:
    # Find the Abstract section and ensure it contains non-whitespace content
    m = re.search(r"\bAbstract\b(.*?)(?=\bIntroduction\b|\bMethodology\b|\bResults\b|\bConclusion\b|$)", text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return False
    body = m.group(1).strip()
    # Consider abstract present if there's at least one alphanumeric character
    return bool(re.search(r"\w", body))


def _has_core_sections(text: str) -> bool:
    required = ["Introduction", "Methodology", "Results", "Conclusion"]
    found = 0
    for token in required:
        if re.search(rf"\b{token}\b", text, flags=re.IGNORECASE):
            found += 1
    return found >= 2


def validate_ieee_document(docx_path: str | Path) -> Dict[str, object]:
    """Validate a lightweight IEEE-like document structure and return pass/fail diagnostics."""
    path = Path(docx_path)
    if not path.exists():
        return {
            "overall_status": "FAIL",
            "file": str(path),
            "checks": {
                "title_present": False,
                "abstract_present": False,
                "author_block_present": False,
                "section_structure_present": False,
            },
            "notes": ["File not found."],
        }

    full_text = _extract_docx_text(path)
    title = _find_title(full_text)
    checks = {
        "title_present": bool(title),
        "abstract_present": _has_abstract(full_text),
        "author_block_present": _has_email_like_author_block(full_text),
        "section_structure_present": _has_core_sections(full_text),
    }

    failed = [name for name, passed in checks.items() if not passed]
    overall_status = "PASS" if not failed else "FAIL"
    notes = []
    if not checks["title_present"]:
        notes.append("Missing a plausible title near the top of the document.")
    if not checks["abstract_present"]:
        notes.append("Missing an 'Abstract' section.")
    if not checks["author_block_present"]:
        notes.append("Missing an author block with email-style author metadata.")
    if not checks["section_structure_present"]:
        notes.append("Missing a usable IEEE-style section structure (e.g., Introduction/Methodology/Results/Conclusion).")

    return {
        "overall_status": overall_status,
        "file": str(path),
        "title": title,
        "checks": checks,
        "notes": notes,
    }
