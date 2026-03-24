"""Integration tests using small in-memory PDF and DOCX files."""

import tempfile
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document

from pdf_docx_to_markdown import convert_document_to_markdown


def _make_simple_pdf(path: Path) -> None:
    """Create a minimal single-page PDF with text."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Hello PDF World", fontsize=12)
    page.insert_text((72, 120), "This is body text.", fontsize=10)
    doc.save(str(path))
    doc.close()


def _make_simple_docx(path: Path) -> None:
    """Create a minimal DOCX with a heading and paragraph."""
    doc = Document()
    doc.add_heading("Test Heading", level=1)
    doc.add_paragraph("This is a test paragraph.")
    doc.save(str(path))


class TestPDFConversion:
    def test_converts_pdf_to_md(self, tmp_path: Path):
        pdf_path = tmp_path / "sample.pdf"
        _make_simple_pdf(pdf_path)

        result = convert_document_to_markdown(pdf_path, tmp_path)

        assert result is not None
        md_path = Path(result)
        assert md_path.exists()
        content = md_path.read_text(encoding="utf-8")
        assert "Hello PDF World" in content
        assert "body text" in content

    def test_pdf_produces_metadata(self, tmp_path: Path):
        pdf_path = tmp_path / "meta.pdf"
        _make_simple_pdf(pdf_path)

        result = convert_document_to_markdown(pdf_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert content.startswith("---")
        assert "source: meta.pdf" in content
        assert "pages:" in content

    def test_missing_file_raises(self, tmp_path: Path):
        import pytest

        with pytest.raises(FileNotFoundError):
            convert_document_to_markdown(tmp_path / "nonexistent.pdf", tmp_path)


class TestDOCXConversion:
    def test_converts_docx_to_md(self, tmp_path: Path):
        docx_path = tmp_path / "sample.docx"
        _make_simple_docx(docx_path)

        result = convert_document_to_markdown(docx_path, tmp_path)

        assert result is not None
        md_path = Path(result)
        assert md_path.exists()
        content = md_path.read_text(encoding="utf-8")
        assert "Test Heading" in content
        assert "test paragraph" in content

    def test_docx_heading_becomes_markdown_heading(self, tmp_path: Path):
        docx_path = tmp_path / "headings.docx"
        _make_simple_docx(docx_path)

        result = convert_document_to_markdown(docx_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert "# Test Heading" in content

    def test_docx_with_table(self, tmp_path: Path):
        docx_path = tmp_path / "table.docx"
        doc = Document()
        doc.add_paragraph("Before table")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Name"
        table.cell(0, 1).text = "Value"
        table.cell(1, 0).text = "foo"
        table.cell(1, 1).text = "bar"
        doc.save(str(docx_path))

        result = convert_document_to_markdown(docx_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert "| Name | Value |" in content
        assert "| foo | bar |" in content
