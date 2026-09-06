"""Tests for helper / utility functions that need no external files."""

from pdf_docx_to_markdown import (
    _blocks_to_markdown,
    _detect_repeating_lines,
    _escape_markdown_cell,
    _extract_text_with_headings,
    _image_extension,
    _table_to_markdown,
    _unique_image_xrefs,
)


class TestEscapeMarkdownCell:
    def test_none_returns_empty(self):
        assert _escape_markdown_cell(None) == ""

    def test_plain_text_unchanged(self):
        assert _escape_markdown_cell("hello world") == "hello world"

    def test_pipe_escaped(self):
        assert _escape_markdown_cell("a|b") == "a\\|b"

    def test_backslash_escaped(self):
        assert _escape_markdown_cell("a\\b") == "a\\\\b"

    def test_newline_becomes_br(self):
        assert _escape_markdown_cell("line1\nline2") == "line1<br>line2"

    def test_crlf_becomes_br(self):
        assert _escape_markdown_cell("line1\r\nline2") == "line1<br>line2"

    def test_whitespace_stripped(self):
        assert _escape_markdown_cell("  spaced  ") == "spaced"

    def test_integer_input(self):
        assert _escape_markdown_cell(42) == "42"


class TestTableToMarkdown:
    def test_empty_table(self):
        assert _table_to_markdown([]) == ""
        assert _table_to_markdown(None) == ""

    def test_header_only(self):
        result = _table_to_markdown([["A", "B"]])
        lines = result.split("\n")
        assert lines[0] == "| A | B |"
        assert lines[1] == "| --- | --- |"
        assert len(lines) == 2

    def test_header_and_rows(self):
        result = _table_to_markdown([["Name", "Age"], ["Alice", "30"], ["Bob", "25"]])
        lines = result.split("\n")
        assert len(lines) == 4
        assert "Alice" in lines[2]
        assert "Bob" in lines[3]

    def test_short_row_padded(self):
        result = _table_to_markdown([["A", "B", "C"], ["x"]])
        lines = result.split("\n")
        assert lines[2].count("|") == lines[0].count("|")

    def test_none_cells_handled(self):
        result = _table_to_markdown([["A", None], [None, "B"]])
        assert "| A |  |" in result
        assert "|  | B |" in result

    def test_special_chars_escaped(self):
        result = _table_to_markdown([["H"], ["a|b"]])
        assert "a\\|b" in result


class TestImageExtension:
    def test_jpeg(self):
        assert _image_extension("image/jpeg") == "jpg"

    def test_png(self):
        assert _image_extension("image/png") == "png"

    def test_svg(self):
        assert _image_extension("image/svg+xml") == "svg"

    def test_emf_skipped(self):
        assert _image_extension("image/x-emf") is None
        assert _image_extension("image/x-wmf") is None

    def test_unknown_skipped(self):
        assert _image_extension("application/octet-stream") is None


class TestBlocksToMarkdown:
    def test_interleaves_by_vertical_position(self):
        blocks = [
            (200.0, "text", "after"),
            (50.0, "text", "before"),
            (100.0, "table", "| A |"),
        ]
        md = _blocks_to_markdown(blocks)
        assert md.index("before") < md.index("| A |") < md.index("after")

    def test_consecutive_text_stays_tight(self):
        md = _blocks_to_markdown(
            [(10.0, "text", "line1"), (20.0, "text", "line2")]
        )
        assert md == "line1\nline2"


class TestPdfLineSort:
    def test_words_sorted_by_x(self):
        class FakePage:
            def extract_words(self, extra_attrs=None):
                return [
                    {"text": "World", "top": 72, "x0": 200, "size": 12},
                    {"text": "Hello", "top": 72, "x0": 72, "size": 12},
                ]

            def extract_text(self):
                return "Hello World"

        assert _extract_text_with_headings(FakePage(), 12) == "Hello World"

    def test_numbered_body_not_promoted(self):
        class FakePage:
            def extract_words(self, extra_attrs=None):
                return [
                    {
                        "text": "1. This is a normal numbered paragraph that should stay body",
                        "top": 72,
                        "x0": 72,
                        "size": 10,
                    }
                ]

            def extract_text(self):
                return ""

        text = _extract_text_with_headings(FakePage(), 10)
        assert not text.startswith("#")
        assert text.startswith("1. ")

    def test_larger_font_is_heading(self):
        class FakePage:
            def extract_words(self, extra_attrs=None):
                return [{"text": "Overview", "top": 72, "x0": 72, "size": 18}]

            def extract_text(self):
                return ""

        assert _extract_text_with_headings(FakePage(), 10) == "# Overview"


class TestRepeatingLines:
    def test_requires_multiple_pages(self):
        assert _detect_repeating_lines([["CONFIDENTIAL"], ["CONFIDENTIAL"]]) == set()

    def test_detects_header(self):
        pages = [["CONFIDENTIAL", "body a"], ["CONFIDENTIAL", "body b"], ["CONFIDENTIAL", "body c"]]
        assert "CONFIDENTIAL" in _detect_repeating_lines(pages)


class TestUniqueXrefs:
    def test_deduplicates(self):
        images = [(7, 0, 40, 40), (7, 0, 40, 40), (9, 0, 40, 40)]
        unique = _unique_image_xrefs(images)
        assert [item[0] for item in unique] == [7, 9]
