"""Tests for the CLI argument parser, batch-mode helpers, and dry-run."""

from pathlib import Path

import fitz

from pdf_docx_to_markdown import (
    ConversionConfig,
    _build_argument_parser,
    _collect_files,
    _should_prompt_between_files,
    main,
)


class TestArgumentParser:
    def test_defaults(self):
        args = _build_argument_parser().parse_args([])
        assert args.file is None
        assert args.output_dir is None
        assert args.recursive is False
        assert args.dry_run is False
        assert args.yes is False
        assert args.verbose is False
        assert args.quiet is False
        assert args.h1_threshold == 6
        assert args.h2_threshold == 3
        assert args.h3_threshold == 2
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

    def test_dry_run_flag(self):
        args = _build_argument_parser().parse_args(["-n"])
        assert args.dry_run is True
        args = _build_argument_parser().parse_args(["--dry-run"])
        assert args.dry_run is True

    def test_yes_flag(self):
        args = _build_argument_parser().parse_args(["-y"])
        assert args.yes is True
        args = _build_argument_parser().parse_args(["--yes"])
        assert args.yes is True

    def test_verbose_and_quiet(self):
        args = _build_argument_parser().parse_args(["-v"])
        assert args.verbose is True
        args = _build_argument_parser().parse_args(["-q"])
        assert args.quiet is True

    def test_custom_thresholds(self):
        args = _build_argument_parser().parse_args(
            [
                "--h1-threshold",
                "10",
                "--h2-threshold",
                "5",
                "--h3-threshold",
                "2",
                "--font-sample-pages",
                "50",
            ]
        )
        assert args.h1_threshold == 10
        assert args.h2_threshold == 5
        assert args.h3_threshold == 2
        assert args.font_sample_pages == 50


class TestCollectFiles:
    def _seed(self, root: Path, recursive: bool = False):
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

    def test_recursive_skips_venv(self, tmp_path: Path):
        venv = tmp_path / ".venv" / "lib"
        venv.mkdir(parents=True)
        (venv / "hidden.pdf").write_bytes(b"fake")
        (tmp_path / "keep.pdf").write_bytes(b"fake")
        names = {f.name for f in _collect_files(tmp_path, recursive=True)}
        assert names == {"keep.pdf"}

    def test_empty_dir(self, tmp_path: Path):
        assert _collect_files(tmp_path) == []


class TestConversionConfig:
    def test_defaults(self):
        cfg = ConversionConfig()
        assert cfg.heading1_threshold == 6
        assert cfg.heading2_threshold == 3
        assert cfg.heading3_threshold == 2
        assert cfg.font_sample_pages == 20
        assert cfg.output_dir is None

    def test_custom(self):
        cfg = ConversionConfig(heading1_threshold=10, font_sample_pages=50)
        assert cfg.heading1_threshold == 10
        assert cfg.font_sample_pages == 50


class TestPrompting:
    def test_yes_disables_prompt(self):
        assert _should_prompt_between_files(True) is False


class TestDryRun:
    def test_dry_run_single_file(self, tmp_path: Path):
        pdf_path = tmp_path / "sample.pdf"
        doc = fitz.open()
        doc.new_page()
        doc.save(str(pdf_path))
        doc.close()

        ret = main([str(pdf_path), "-o", str(tmp_path), "-n"])
        assert ret == 0
        assert not list(tmp_path.glob("*.md"))

    def test_dry_run_batch(self, tmp_path: Path, monkeypatch):
        (tmp_path / "a.pdf").write_bytes(b"%PDF-fake")
        (tmp_path / "b.docx").write_bytes(b"fake")
        monkeypatch.chdir(tmp_path)

        ret = main(["-n"])
        assert ret == 0
        assert not (tmp_path / "md_output").exists()


class TestBatchCwdAndYes:
    def _pdf(self, path: Path) -> None:
        doc = fitz.open()
        page = doc.new_page()
        page.insert_text((72, 72), "Hello", fontsize=12)
        doc.save(str(path))
        doc.close()

    def test_batch_uses_cwd(self, tmp_path: Path, monkeypatch):
        self._pdf(tmp_path / "a.pdf")
        monkeypatch.chdir(tmp_path)
        out = tmp_path / "out"
        ret = main(["-y", "-o", str(out)])
        assert ret == 0
        assert (out / "a.md").exists()

    def test_yes_skips_prompt(self, tmp_path: Path, monkeypatch):
        self._pdf(tmp_path / "a.pdf")
        self._pdf(tmp_path / "b.pdf")
        monkeypatch.chdir(tmp_path)

        def boom(*_args, **_kwargs):
            raise AssertionError("input should not be called")

        monkeypatch.setattr("builtins.input", boom)
        ret = main(["-y", "-o", str(tmp_path / "out")])
        assert ret == 0
        assert (tmp_path / "out" / "a.md").exists()
        assert (tmp_path / "out" / "b.md").exists()

    def test_non_tty_skips_prompt(self, tmp_path: Path, monkeypatch):
        self._pdf(tmp_path / "a.pdf")
        self._pdf(tmp_path / "b.pdf")
        monkeypatch.chdir(tmp_path)

        class FakeStdin:
            def isatty(self):
                return False

        monkeypatch.setattr("pdf_docx_to_markdown.sys.stdin", FakeStdin())
        monkeypatch.setattr(
            "builtins.input",
            lambda *_a, **_k: (_ for _ in ()).throw(AssertionError("input called")),
        )
        ret = main(["-o", str(tmp_path / "out")])
        assert ret == 0
        assert (tmp_path / "out" / "a.md").exists()
        assert (tmp_path / "out" / "b.md").exists()
