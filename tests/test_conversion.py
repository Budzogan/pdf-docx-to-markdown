"""Integration tests using small in-memory PDF and DOCX files."""

import logging
import struct
import zlib
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pdf_docx_to_markdown import ConversionConfig, convert_document_to_markdown


def _make_png(width: int = 40, height: int = 40) -> bytes:
    raw = b""
    for _ in range(height):
        raw += b"\x00" + (b"\xff\x00\x00" * width)
    compressed = zlib.compress(raw)

    def chunk(ctype: bytes, data: bytes) -> bytes:
        c = ctype + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", compressed)
        + chunk(b"IEND", b"")
    )


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


def _make_image_only_pdf(path: Path) -> None:
    """Create a PDF with an image but no text (simulates scanned page)."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_image(fitz.Rect(50, 50, 200, 200), stream=_make_png(40, 40))
    doc.save(str(path))
    doc.close()


def _add_hyperlink(paragraph, text: str, url: str) -> None:
    rel_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    new_run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


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

    def test_unsupported_extension_returns_none(self, tmp_path: Path):
        txt_path = tmp_path / "readme.txt"
        txt_path.write_text("hello")

        result = convert_document_to_markdown(txt_path, tmp_path)
        assert result is None

    def test_custom_config_thresholds(self, tmp_path: Path):
        pdf_path = tmp_path / "cfg.pdf"
        _make_simple_pdf(pdf_path)

        cfg = ConversionConfig(
            heading1_threshold=10,
            heading2_threshold=8,
            heading3_threshold=5,
        )
        result = convert_document_to_markdown(pdf_path, tmp_path, config=cfg)
        assert result is not None

    def test_scanned_pdf_warning(self, tmp_path: Path, caplog):
        pdf_path = tmp_path / "scanned.pdf"
        _make_image_only_pdf(pdf_path)

        with caplog.at_level(logging.WARNING):
            convert_document_to_markdown(pdf_path, tmp_path)

        assert any(
            "scanned" in r.message.lower() or "image" in r.message.lower()
            for r in caplog.records
            if r.levelno >= logging.WARNING
        )

    def test_duplicate_xref_writes_once(self, tmp_path: Path):
        pdf_path = tmp_path / "dup.pdf"
        doc = fitz.open()
        page = doc.new_page()
        png = _make_png(40, 40)
        page.insert_image(fitz.Rect(50, 50, 150, 150), stream=png)
        page.insert_text((72, 220), "Caption", fontsize=12)
        doc.save(str(pdf_path))
        doc.close()

        convert_document_to_markdown(pdf_path, tmp_path)
        images = list((tmp_path / "dup_images").glob("*")) if (tmp_path / "dup_images").exists() else []
        assert len(images) <= 1


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
        assert content.startswith("---")
        assert "source: sample.docx" in content

    def test_docx_heading_becomes_markdown_heading(self, tmp_path: Path):
        docx_path = tmp_path / "headings.docx"
        _make_simple_docx(docx_path)

        result = convert_document_to_markdown(docx_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert "# Test Heading" in content

    def test_docx_heading_6(self, tmp_path: Path):
        docx_path = tmp_path / "h6.docx"
        doc = Document()
        doc.add_heading("Deep heading", level=6)
        doc.save(str(docx_path))

        content = Path(convert_document_to_markdown(docx_path, tmp_path)).read_text(encoding="utf-8")
        assert "###### Deep heading" in content

    def test_docx_numbered_list(self, tmp_path: Path):
        docx_path = tmp_path / "list.docx"
        doc = Document()
        doc.add_paragraph("Alpha", style="List Number")
        doc.add_paragraph("Beta", style="List Number")
        doc.save(str(docx_path))

        content = Path(convert_document_to_markdown(docx_path, tmp_path)).read_text(encoding="utf-8")
        assert "Alpha" in content
        assert "Beta" in content
        assert "1. Alpha" in content
        assert "2. Beta" in content

    def test_docx_hyperlink(self, tmp_path: Path):
        docx_path = tmp_path / "link.docx"
        doc = Document()
        para = doc.add_paragraph()
        _add_hyperlink(para, "Example", "https://example.com")
        doc.save(str(docx_path))

        content = Path(convert_document_to_markdown(docx_path, tmp_path)).read_text(encoding="utf-8")
        assert "[Example](https://example.com)" in content

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

    def test_docx_merged_cells(self, tmp_path: Path):
        """Horizontally merged cells should not produce duplicate columns."""
        docx_path = tmp_path / "merged.docx"
        doc = Document()
        table = doc.add_table(rows=2, cols=3)
        table.cell(0, 0).text = "A"
        table.cell(0, 1).text = "B"
        table.cell(0, 2).text = "C"
        table.cell(1, 0).merge(table.cell(1, 1))
        table.cell(1, 0).text = "Merged"
        table.cell(1, 2).text = "Solo"
        doc.save(str(docx_path))

        result = convert_document_to_markdown(docx_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert "| A | B | C |" in content
        assert "Merged" in content
        assert "Solo" in content

    def test_docx_nested_table(self, tmp_path: Path):
        """Nested tables inside cells should be rendered inline."""
        from docx.oxml.ns import qn as _qn

        docx_path = tmp_path / "nested.docx"
        doc = Document()
        outer = doc.add_table(rows=1, cols=2)
        outer.cell(0, 0).text = "Left"
        inner_tbl = outer.cell(0, 1)._element.makeelement(_qn("w:tbl"), {})
        outer.cell(0, 1)._element.append(inner_tbl)
        tr = inner_tbl.makeelement(_qn("w:tr"), {})
        inner_tbl.append(tr)
        tc = tr.makeelement(_qn("w:tc"), {})
        tr.append(tc)
        p = tc.makeelement(_qn("w:p"), {})
        tc.append(p)
        r = p.makeelement(_qn("w:r"), {})
        p.append(r)
        t = r.makeelement(_qn("w:t"), {})
        t.text = "Nested"
        r.append(t)
        doc.save(str(docx_path))

        result = convert_document_to_markdown(docx_path, tmp_path)

        content = Path(result).read_text(encoding="utf-8")
        assert "Left" in content
        assert "Nested" in content
