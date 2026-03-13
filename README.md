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

---

## What it does

- Converts `.pdf` and `.docx` files to `.md`
- Extracts and saves images from PDFs
- Detects headings, tables, and multi-column layouts
- Processes multiple files in one run with a `y/n` prompt between each
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

| Package | Purpose |
|---|---|
| `PyMuPDF` | Extracts images from PDFs |
| `pdfplumber` | Extracts text and tables from PDFs |
| `python-docx` | Converts Word `.docx` files without AI dependencies |
| `tqdm` | Displays a terminal progress bar |

> Note: Total dependency download is under 50 MB.

If you use `CONVERT_DOCS.bat`, you usually do not need to do this manually. The batch file tries to install what is missing for you.

---

## Usage

### Option A - Double-click `CONVERT_DOCS.bat` (recommended for most people)

1. Put your `.pdf` or `.docx` files in this folder.
2. Double-click `CONVERT_DOCS.bat`.
3. The output appears in the `md_output\` folder.

What the batch file does:

1. Checks whether Python is already installed.
2. If Python is missing, it tries to install it automatically.
3. Installs the required Python libraries.
4. Runs the converter.
5. Opens the output folder when finished.

On the first run, setup can take a few minutes.

### Option B - Command line

Convert all files in the folder:

```bash
python pdf_docx_to_markdown.py
```

Convert a specific file:

```bash
python pdf_docx_to_markdown.py "path/to/file.pdf"
```

Convert a specific file to a custom output folder:

```bash
python pdf_docx_to_markdown.py "path/to/file.pdf" "path/to/output/"
```

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

  Next: annex.docx  (1 file(s) remaining)
  Continue? [y/n]:
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

- **PDF**: `pdfplumber` extracts text and tables, while `PyMuPDF` extracts images. Headings are detected by comparing font sizes to the body text size.
- **DOCX**: `python-docx` reads the Word XML directly. Heading styles, tables, lists, and embedded images are extracted without AI models.

---

## Known Limitations

- PDF conversion quality depends heavily on how well the original PDF is structured
- Very complex layouts can still produce imperfect reading order
- Scientific papers, equations, references, and multi-layer academic formatting are not the main target
- Scanned PDFs without usable text layers may need OCR-focused tools instead
- Markdown tables are simplified representations of the original table layout

This tool aims to be practical, lightweight, and local-first. It is not trying to be a perfect document reconstruction engine.

---

## Notes

- Progress bars use ASCII characters so they display more reliably in Windows terminals.
- The batch file uses `python -m pip`, so the common `pip.exe is not on PATH` warning is not a blocker for normal use.
- `CONVERT_DOCS.bat` is the easiest option for non-technical Windows users.

---

## Files

| File | Purpose |
|---|---|
| `pdf_docx_to_markdown.py` | Main conversion script |
| `requirements_extract.txt` | Python dependencies |
| `CONVERT_DOCS.bat` | One-click runner for Windows |
| `README.md` | Project documentation |

---

## License

MIT

---

## Support

Commercial users who find value in this tool are encouraged to sponsor the project.

Enjoy.
