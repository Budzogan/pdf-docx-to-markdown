# pdf-docx-to-markdown

Convert PDF and DOCX files to clean Markdown locally, with no cloud services or API keys.

Designed to produce LLM-ready output with readable structure, preserved images, and usable Markdown for notes, internal docs, specs, and document cleanup.

---

## Quick Start For Windows

If you just want to use it and do not care about the technical details:

1. Put your `.pdf` or `.docx` files in this folder.
2. Double-click `CONVERT_DOCS.bat`.
3. Wait while it checks Python and installs anything missing.
4. If Windows asks for permission to continue, allow it.
5. Your Markdown files will appear in the `md_output\` folder.

That is the main way this tool is meant to be used on Windows.

The batch file creates a local `.venv` in this folder (not a global Python install) and converts every PDF/DOCX in the folder without asking `y/n` between files.

---

## What it does

- Converts `.pdf` and `.docx` files to `.md`
- Extracts and saves images from PDFs and DOCX files (skips tiny decorative images and EMF/WMF)
- Detects headings and tables (note: multi-column PDFs are not fully supported and may produce merged text)
- Processes multiple files in one run (`CONVERT_DOCS.bat` and `-y` run without prompts)
- Shows a progress bar and timing in the terminal for each file
- Runs fully offline after first install

---

## Who This Is For

- People who want a simple local converter without cloud uploads
- Non-technical Windows users who prefer double-clicking a `.bat` file
- Users cleaning up business documents, product specs, manuals, reports, and general office files
- Anyone who wants Markdown output to reuse in ChatGPT, Copilot, Claude, or other LLM tools

## Who This Is Not For

- Users expecting perfect reconstruction of every PDF layout
- Academic or research-heavy workflows with dense formulas, citations, footnotes, or complex paper structure
- Teams looking for OCR pipelines, enterprise document ingestion, or high-accuracy scientific parsing

If your main goal is converting scientific papers or complex academic PDFs, a more specialized tool such as a Docling-based converter may be a better fit.

---

## Prerequisites

Most Windows users do not need to install anything manually. If you use `CONVERT_DOCS.bat`, the tool will try to set up Python 3.10+ and the required libraries for you automatically.

The manual steps below are mainly for people who want to run the script from the command line or fix a local Python setup themselves.

### 1. Install Python 3.10 or newer

Download from [python.org](https://www.python.org/downloads/).

> Warning: During install, check **"Add Python to PATH"**

Verify it works:

```bash
python --version
```

### 2. Install required libraries

Open a terminal in this folder and run:

```bash
pip install -r requirements_extract.txt
```

Or, from a clone:

```bash
pip install -e ".[dev]"
```

| Package | Purpose |
|---|---|
| `PyMuPDF` | Extracts images from PDFs |
| `pdfplumber` | Extracts text and tables from PDFs |
| `python-docx` | Converts Word `.docx` files without AI dependencies |
| `tqdm` | Displays a terminal progress bar |

> Note: Total dependency download is under 50 MB.

---

## Usage

### Option A - Double-click `CONVERT_DOCS.bat` (recommended for most people)

1. Put your `.pdf` or `.docx` files in this folder.
2. Double-click `CONVERT_DOCS.bat`.
3. The output appears in the `md_output\` folder.

What the batch file does:

1. Checks for Python 3.10 or newer (`py -3`, then `python`).
2. If Python is missing, it tries to install it automatically.
3. Creates a local `.venv` and installs libraries only if they are missing.
4. Runs the converter with `-y` (no prompt between files).
5. Opens the output folder when finished.

On the first run, setup can take a few minutes.

### Option B - Command line

Convert all PDF/DOCX files in the **current directory**:

```bash
python pdf_docx_to_markdown.py -y
```

Convert a specific file:

```bash
python pdf_docx_to_markdown.py "path/to/file.pdf"
```

Convert a specific file to a custom output folder:

```bash
python pdf_docx_to_markdown.py "path/to/file.pdf" -o "path/to/output/"
```

| Flag | Meaning |
|---|---|
| `-o`, `--output-dir` | Output folder (default: `./md_output/`) |
| `-r`, `--recursive` | In batch mode, include subfolders |
| `-n`, `--dry-run` | List files that would be converted |
| `-y`, `--yes` | Do not prompt between files |
| `-v`, `--verbose` | Debug logging |
| `-q`, `--quiet` | Errors only |
| `--h1-threshold`, `--h2-threshold`, `--h3-threshold` | PDF heading font-size gaps (pt) |
| `--font-sample-pages` | Pages sampled to detect body font size |
| `--version` | Print version |

Without `-y`, batch mode asks `Continue? [Y/n]` between files (Enter = yes). The prompt is skipped when stdin is not a terminal.

---

## What You Get

After conversion, you get:

- A `.md` file for each source document
- A separate images folder when the source contains extracted images
- Markdown that is usually easier to read, edit, search, and paste into LLM tools than the original document

---

## Terminal output example

```text
============================================================
  File    : specification.pdf
  Size    : 4.23 MB
  Started : 14:32:05
============================================================
  [1/2] Extracting images: 100%|##########| 87/87 [00:04<00:00, 21pg/s]
  [2/2] Extracting text : 100%|##########| 87/87 [00:18<00:00,  4pg/s]

============================================================
  DONE!
  Output  : specification.md
  Size    : 312.4 KB
  Time    : 22s
  Finished: 14:32:27
============================================================
```

---

## Output structure

```text
md_output/
|-- specification.md
|-- specification_images/
|   |-- page1_img1.png
|   `-- page3_img1.png
`-- annex.md
```

---

## How it works

- **PDF**: `pdfplumber` extracts text and tables, while `PyMuPDF` extracts images. Headings are detected by comparing font sizes to the body text size. Text, tables, and images are interleaved by vertical position on the page.
- **DOCX**: `python-docx` reads the Word XML directly. Heading styles (including Heading 6), tables, numbered/bulleted lists, hyperlinks, and embedded images are extracted without AI models.

---

## Known Limitations

- PDF conversion quality depends heavily on how well the original PDF is structured
- Multi-column PDFs are not fully supported and may produce merged text
- Very complex layouts can still produce imperfect reading order
- Scientific papers, equations, references, and multi-layer academic formatting are not the main target
- Scanned PDFs without usable text layers may need OCR-focused tools instead
- Markdown tables are simplified representations of the original table layout

This tool aims to be practical, lightweight, and local-first. It is not trying to be a perfect document reconstruction engine.

---

## Notes

- Progress bars use ASCII characters so they display more reliably in Windows terminals.
- The batch file uses `python -m pip` inside `.venv`, so the common `pip.exe is not on PATH` warning is not a blocker for normal use.
- `CONVERT_DOCS.bat` is the easiest option for non-technical Windows users.

---

## Files

| File | Purpose |
|---|---|
| `pdf_docx_to_markdown.py` | Main conversion script |
| `pyproject.toml` | Package metadata and dependencies |
| `requirements_extract.txt` | Python dependencies for the batch installer |
| `CONVERT_DOCS.bat` | One-click runner for Windows |
| `tests/` | pytest suite |
| `README.md` | Project documentation |
| `LICENSE` | MIT license |

---

## License

MIT

---

## Support / Contact

- Bug reports and feature requests: open a GitHub Issue
- Questions and ideas: use GitHub Discussions
