"""Tests for the CLI argument parser and batch-mode helpers."""

from pathlib import Path

from pdf_docx_to_markdown import (
    _build_argument_parser,
    _collect_files,
    ConversionConfig,
)


class TestArgumentParser:
    def test_defaults(self):
        args = _build_argument_parser().parse_args([])
        assert args.file is None
        assert args.output_dir is None
        assert args.recursive is False
        assert args.verbose is False
        assert args.quiet is False
        assert args.h1_threshold == 6
        assert args.h2_threshold == 3
        assert args.h3_threshold == 1
        assert args.font_sample_pages == 20

    def test_single_file(self):
        args = _build_argument_parser().parse_args(["my.pdf"])
        assert args.file == "my.pdf"

    def test_output_dir_short_flag(self):
        args = _build_argument_parser().parse_args(["-o", "/tmp/out"])
        assert args.output_dir == "/tmp/out"

    def test_recursive_flag(self):
        args = _build_argument_parser().parse_args(["-r"])
        assert args.recursive is True

    def test_verbose_and_quiet(self):
        args = _build_argument_parser().parse_args(["-v"])
        assert args.verbose is True
        args = _build_argument_parser().parse_args(["-q"])
        assert args.quiet is True

    def test_custom_thresholds(self):
        args = _build_argument_parser().parse_args([
            "--h1-threshold", "10",
            "--h2-threshold", "5",
            "--h3-threshold", "2",
            "--font-sample-pages", "50",
        ])
        assert args.h1_threshold == 10
        assert args.h2_threshold == 5
        assert args.h3_threshold == 2
        assert args.font_sample_pages == 50


class TestCollectFiles:
    def _seed(self, root: Path, recursive: bool = False):
        """Create test files in root (and optionally a subdirectory)."""
        (root / "a.pdf").write_bytes(b"fake")
        (root / "b.docx").write_bytes(b"fake")
        (root / "c.txt").write_bytes(b"ignore")
        if recursive:
            sub = root / "sub"
            sub.mkdir()
            (sub / "d.pdf").write_bytes(b"fake")

    def test_flat(self, tmp_path: Path):
        self._seed(tmp_path)
        files = _collect_files(tmp_path, recursive=False)
        names = {f.name for f in files}
        assert names == {"a.pdf", "b.docx"}

    def test_recursive(self, tmp_path: Path):
        self._seed(tmp_path, recursive=True)
        files = _collect_files(tmp_path, recursive=True)
        names = {f.name for f in files}
        assert names == {"a.pdf", "b.docx", "d.pdf"}

    def test_flat_skips_subdirs(self, tmp_path: Path):
        self._seed(tmp_path, recursive=True)
        files = _collect_files(tmp_path, recursive=False)
        names = {f.name for f in files}
        assert "d.pdf" not in names

    def test_empty_dir(self, tmp_path: Path):
        assert _collect_files(tmp_path) == []


class TestConversionConfig:
    def test_defaults(self):
        cfg = ConversionConfig()
        assert cfg.heading1_threshold == 6
        assert cfg.heading2_threshold == 3
        assert cfg.heading3_threshold == 1
        assert cfg.font_sample_pages == 20
        assert cfg.output_dir is None

    def test_custom(self):
        cfg = ConversionConfig(heading1_threshold=10, font_sample_pages=50)
        assert cfg.heading1_threshold == 10
        assert cfg.font_sample_pages == 50
