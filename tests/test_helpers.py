"""Tests for helper / utility functions that need no external files."""

from pdf_docx_to_markdown import _escape_markdown_cell, _table_to_markdown


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
        # Short row should be padded to match header width
        assert lines[2].count("|") == lines[0].count("|")

    def test_none_cells_handled(self):
        result = _table_to_markdown([["A", None], [None, "B"]])
        assert "| A |  |" in result
        assert "|  | B |" in result

    def test_special_chars_escaped(self):
        result = _table_to_markdown([["H"], ["a|b"]])
        assert "a\\|b" in result
