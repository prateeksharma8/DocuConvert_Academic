import os
import tempfile
import unittest
import zipfile
from pathlib import Path

from services.ieee_validator import validate_ieee_document


def build_docx(path: Path, title: str, abstract: str, sections: str):
    body = f"""
    <w:p><w:r><w:t>{title}</w:t></w:r></w:p>
    <w:p><w:r><w:t>1 Author One</w:t></w:r></w:p>
    <w:p><w:r><w:t>2 Author Two</w:t></w:r></w:p>
    <w:p><w:r><w:t>Abstract</w:t></w:r></w:p>
    <w:p><w:r><w:t>{abstract}</w:t></w:r></w:p>
    <w:p><w:r><w:t>Introduction</w:t></w:r></w:p>
    <w:p><w:r><w:t>{sections}</w:t></w:r></w:p>
    <w:p><w:r><w:t>Methodology</w:t></w:r></w:p>
    <w:p><w:r><w:t>Results</w:t></w:r></w:p>
    <w:p><w:r><w:t>Conclusion</w:t></w:r></w:p>
    <w:p><w:r><w:t>References</w:t></w:r></w:p>
    """.strip()
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>
</Types>""")
        zf.writestr("_rels/.rels", """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>
</Relationships>""")
        zf.writestr("word/document.xml", f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:body>
    {body}
    <w:p/>
  </w:body>
</w:document>""")


class TestIEEEValidator(unittest.TestCase):
    def test_reference_style_document_passes(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = Path(tmpdir) / "sample.docx"
            build_docx(
                docx_path,
                "A Lightweight Validation Framework for IEEE-Style Technical Documents",
                "This abstract explains the purpose of the document and highlights the validation logic.",
                "The paper presents a lightweight validation framework for IEEE-style documents.",
            )

            result = validate_ieee_document(docx_path)
            self.assertEqual(result["overall_status"], "PASS")
            self.assertTrue(result["checks"]["title_present"])
            self.assertTrue(result["checks"]["abstract_present"])
            self.assertTrue(result["checks"]["author_block_present"])
            self.assertTrue(result["checks"]["section_structure_present"])

    def test_missing_abstract_fails(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = Path(tmpdir) / "bad.docx"
            build_docx(
                docx_path,
                "A Lightweight Validation Framework for IEEE-Style Technical Documents",
                "",
                "The paper presents a validation framework.",
            )

            result = validate_ieee_document(docx_path)
            self.assertEqual(result["overall_status"], "FAIL")
            self.assertFalse(result["checks"]["abstract_present"])


if __name__ == "__main__":
    unittest.main()
