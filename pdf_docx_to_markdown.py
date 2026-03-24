import argparse
import logging
import re
import sys
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from tqdm import tqdm

import fitz  # PyMuPDF - for image extraction from PDFs
import pdfplumber  # For text + table extraction from PDFs
from docx import Document  # python-docx - for DOCX conversion
from docx.oxml.ns import qn

SUPPORTED_EXTENSIONS = {".docx", ".pdf"}
SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "md_output"

logger = logging.getLogger(__name__)


@dataclass
class ConversionConfig:
    """Configurable thresholds and limits for document conversion."""

    # Font-size difference thresholds for heading detection (PDF only).
    heading1_threshold: int = 6
    heading2_threshold: int = 3
    heading3_threshold: int = 1

    # Maximum pages sampled when detecting the body font size.
    font_sample_pages: int = 20

    # Output directory (None = same directory as source file).
    output_dir: Path | None = None


# Singleton default used when callers don't pass an explicit config.
_DEFAULT_CONFIG = ConversionConfig()


def convert_document_to_markdown(
    input_path: str | Path,
    output_dir: str | Path | None = None,
    config: ConversionConfig | None = None,
) -> str | None:
    cfg = config or _DEFAULT_CONFIG
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"File not found: {source}")

    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        logger.error(
            "Unsupported file type '%s'. Supported: %s",
            source.suffix, ", ".join(sorted(SUPPORTED_EXTENSIONS)),
        )
        return None

    effective_output = (
        Path(output_dir).expanduser().resolve()
        if output_dir
        else (cfg.output_dir.expanduser().resolve() if cfg.output_dir else source.parent)
    )
    effective_output.mkdir(parents=True, exist_ok=True)

    file_size_mb = source.stat().st_size / (1024 * 1024)
    logger.info(
        "\n%s\n  File    : %s\n  Size    : %.2f MB\n  Started : %s\n%s",
        "=" * 60, source.name, file_size_mb,
        datetime.now().strftime("%H:%M:%S"), "=" * 60,
    )

    t_start = time.time()

    if source.suffix.lower() == ".pdf":
        markdown_content = _convert_pdf(source, effective_output, cfg)
    else:
        logger.info("  Processing DOCX...")
        markdown_content = _convert_with_python_docx(source, effective_output)

    elapsed = time.time() - t_start

    if markdown_content is None:
        logger.error("  ERROR: Failed to convert %s.", source.name)
        return None

    md_path = effective_output / f"{source.stem}.md"
    md_path.write_text(markdown_content, encoding="utf-8")

    out_size_kb = md_path.stat().st_size / 1024
    mins, secs = divmod(int(elapsed), 60)
    time_str = f"{mins}m {secs}s" if mins else f"{secs:.1f}s"

    logger.info(
        "\n%s\n  DONE!\n  Output  : %s\n  Size    : %.1f KB\n  Time    : %s\n  Finished: %s\n%s\n",
        "=" * 60, md_path.name, out_size_kb, time_str,
        datetime.now().strftime("%H:%M:%S"), "=" * 60,
    )
    return str(md_path)


def _convert_with_python_docx(source: Path, target_dir: Path) -> str | None:
    """Convert DOCX using python-docx - lightweight, no AI/ML dependencies."""
    try:
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        doc = Document(str(source))
        md_lines: list[str] = []

        heading_map: dict[str, str] = {
            "Title": "# ",
            "Subtitle": "## ",
            "Heading 1": "# ",
            "Heading 2": "## ",
            "Heading 3": "### ",
            "Heading 4": "#### ",
            "Heading 5": "##### ",
        }

        images_dir = target_dir / f"{source.stem}_images"
        image_map: dict[str, str] = {}
        try:
            for rel in doc.part.rels.values():
                if "image" not in rel.reltype:
                    continue
                img_data = rel.target_part.blob
                ext = rel.target_part.content_type.split("/")[-1]
                img_filename = f"img_{len(image_map) + 1}.{ext}"
                images_dir.mkdir(parents=True, exist_ok=True)
                (images_dir / img_filename).write_bytes(img_data)
                image_map[rel.rId] = img_filename
        except Exception:
            pass

        elements = list(doc.element.body)
        for element in tqdm(
            elements,
            desc="  Processing elements",
            unit="el",
            ncols=60,
            ascii=True,
            disable=len(elements) < 50,
        ):
            tag = element.tag.split("}")[-1]

            if tag == "p":
                para = Paragraph(element, doc)
                content = _extract_docx_paragraph_content(para, source.stem, image_map).strip()

                if not content:
                    md_lines.append("")
                    continue

                style_name = para.style.name if para.style else ""
                prefix = heading_map.get(style_name, "")
                if prefix:
                    md_lines.append(f"\n{prefix}{content}\n")
                    continue

                num_pr = element.find(".//" + qn("w:numPr"))
                if num_pr is not None:
                    ilvl = num_pr.find(qn("w:ilvl"))
                    level = int(ilvl.get(qn("w:val"), 0)) if ilvl is not None else 0
                    md_lines.append("  " * level + f"- {content}")
                else:
                    md_lines.append(content)

            elif tag == "tbl":
                table = Table(element, doc)
                rows = table.rows
                if not rows:
                    continue

                header = [_escape_markdown_cell(cell.text) for cell in rows[0].cells]
                md_lines.append("")
                md_lines.append("| " + " | ".join(header) + " |")
                md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")

                for row in rows[1:]:
                    cells = [_escape_markdown_cell(cell.text) for cell in row.cells]
                    md_lines.append("| " + " | ".join(cells) + " |")

                md_lines.append("")

        if image_map:
            logger.info("  Extracted %d image(s) to %s", len(image_map), images_dir)
        elif images_dir.exists():
            try:
                images_dir.rmdir()
            except OSError:
                pass

        return "\n".join(md_lines)
    except Exception as e:
        logger.error("python-docx conversion failed: %s", e)
        return None


def _convert_pdf(
    source: Path, target_dir: Path, cfg: ConversionConfig
) -> str | None:
    """Extract text, tables, and images from PDF using pdfplumber + PyMuPDF."""
    try:
        images_dir = target_dir / f"{source.stem}_images"
        images_dir.mkdir(parents=True, exist_ok=True)

        md_parts: list[str] = []
        image_count = 0
        body_font_size = _detect_body_font_size(source, cfg.font_sample_pages)

        fitz_doc = fitz.open(str(source))
        total_pages = len(fitz_doc)
        page_images: dict[int, list[str]] = {}

        for page_num in tqdm(
            range(total_pages),
            desc="  [1/2] Extracting images",
            unit="pg",
            ncols=60,
            ascii=True,
        ):
            page = fitz_doc[page_num]
            image_list = page.get_images(full=True)
            page_imgs: list[str] = []
            for img_idx, img_info in enumerate(image_list):
                xref = img_info[0]
                try:
                    base_image = fitz_doc.extract_image(xref)
                    if base_image and base_image["image"]:
                        ext = base_image.get("ext", "png")
                        img_filename = f"page{page_num + 1}_img{img_idx + 1}.{ext}"
                        img_path = images_dir / img_filename
                        img_path.write_bytes(base_image["image"])
                        page_imgs.append(img_filename)
                        image_count += 1
                except Exception:
                    pass
            page_images[page_num] = page_imgs
        fitz_doc.close()

        # Track pages that look scanned (images but no text).
        scanned_page_count = 0

        with pdfplumber.open(str(source)) as pdf:
            for page_num, page in tqdm(
                enumerate(pdf.pages),
                desc="  [2/2] Extracting text ",
                unit="pg",
                ncols=60,
                total=total_pages,
                ascii=True,
            ):
                page_md = f"<!-- Page {page_num + 1} -->\n\n"

                tables = page.extract_tables()
                table_bboxes: list[tuple[float, float, float, float]] = []
                if tables:
                    for table in page.find_tables():
                        table_bboxes.append(table.bbox)

                if table_bboxes:
                    filtered_page = page
                    for bbox in table_bboxes:
                        filtered_page = filtered_page.outside_bbox(bbox)
                else:
                    filtered_page = page

                text = _extract_text_with_headings(
                    filtered_page, body_font_size, cfg
                )
                has_text = bool(text.strip()) or bool(tables)
                if text.strip():
                    page_md += text.strip() + "\n\n"

                for table in tables:
                    if table and len(table) > 0:
                        page_md += _table_to_markdown(table) + "\n\n"

                page_has_images = bool(page_images.get(page_num))
                if page_has_images and not has_text:
                    scanned_page_count += 1

                for img_file in page_images.get(page_num, []):
                    rel_path = f"{source.stem}_images/{img_file}"
                    page_md += f"![Image]({rel_path})\n\n"

                if page_md.strip() != f"<!-- Page {page_num + 1} -->":
                    md_parts.append(page_md.rstrip())

        if image_count > 0:
            logger.info("  Extracted %d image(s) to %s", image_count, images_dir)
        else:
            try:
                images_dir.rmdir()
            except OSError:
                pass

        # Warn about likely scanned pages.
        if scanned_page_count > 0:
            pct = scanned_page_count / total_pages * 100
            logger.warning(
                "  WARNING: %d of %d page(s) (%.0f%%) appear to be scanned "
                "(images found but no extractable text).",
                scanned_page_count, total_pages, pct,
            )
            if scanned_page_count == total_pages:
                logger.warning(
                    "  This PDF seems to be fully scanned / image-based. "
                    "Consider running an OCR tool (e.g. Tesseract, Adobe Acrobat) "
                    "to add a text layer before converting."
                )

        if not md_parts:
            logger.warning("No content extracted from %s", source.name)
            return ""

        metadata = (
            f"---\n"
            f"source: {source.name}\n"
            f"pages: {total_pages}\n"
            f"extracted: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"---\n\n"
        )

        return metadata + "\n\n---\n\n".join(md_parts)
    except Exception as e:
        logger.error("PDF extraction failed: %s", e)
        return None


def _detect_body_font_size(source: Path, sample_pages: int = 20) -> int:
    """Scan the whole PDF and return the most common font size (= body text)."""
    try:
        all_sizes: list[int] = []
        with pdfplumber.open(str(source)) as pdf:
            for page in pdf.pages[:sample_pages]:
                words = page.extract_words(extra_attrs=["size"])
                all_sizes.extend(round(word["size"]) for word in words if word.get("size"))
        if not all_sizes:
            return 10
        return Counter(all_sizes).most_common(1)[0][0]
    except Exception:
        return 10


def _extract_text_with_headings(
    page: object, body_font_size: int, cfg: ConversionConfig | None = None
) -> str:
    """Extract page text, promoting heading lines to markdown # notation."""
    if cfg is None:
        cfg = _DEFAULT_CONFIG

    try:
        words = page.extract_words(extra_attrs=["size"])
    except Exception:
        return page.extract_text() or ""

    if not words:
        return ""

    line_buckets: dict[int, list[dict]] = {}
    for word in words:
        top = round(word["top"] / 3) * 3
        line_buckets.setdefault(top, []).append(word)

    result_lines: list[str] = []
    for top in sorted(line_buckets):
        line_words = line_buckets[top]
        text = " ".join(word["text"] for word in line_words).strip()
        if not text:
            continue

        sizes = [word["size"] for word in line_words if word.get("size")]
        avg_size = round(sum(sizes) / len(sizes)) if sizes else body_font_size
        diff = avg_size - body_font_size

        if diff >= cfg.heading1_threshold:
            text = f"# {text}"
        elif diff >= cfg.heading2_threshold:
            text = f"## {text}"
        elif diff >= cfg.heading3_threshold:
            text = f"### {text}"
        else:
            if re.match(r"^\d+\.\d+\.\d+[\s\.]", text) and len(text) < 120:
                text = f"#### {text}"
            elif re.match(r"^\d+\.\d+[\s\.]", text) and len(text) < 120:
                text = f"### {text}"
            elif re.match(r"^\d+\.\s+\S", text) and len(text) < 120:
                text = f"## {text}"

        result_lines.append(text)

    return "\n".join(result_lines)


def _table_to_markdown(table: list[list[str | None]]) -> str:
    """Convert a pdfplumber table (list of lists) to markdown table format."""
    if not table:
        return ""

    clean: list[list[str]] = []
    for row in table:
        clean.append([_escape_markdown_cell(cell) for cell in row])

    if not clean:
        return ""

    lines: list[str] = []
    lines.append("| " + " | ".join(clean[0]) + " |")
    lines.append("| " + " | ".join(["---"] * len(clean[0])) + " |")

    for row in clean[1:]:
        while len(row) < len(clean[0]):
            row.append("")
        lines.append("| " + " | ".join(row[: len(clean[0])]) + " |")

    return "\n".join(lines)


def _extract_docx_paragraph_content(
    para: object, source_stem: str, image_map: dict[str, str]
) -> str:
    """Return paragraph content with images inserted in run order."""
    parts: list[str] = []

    for run in para.runs:
        if run.text:
            parts.append(run.text)

        for blip in run._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            rel_id = blip.get(qn("r:embed"))
            img_filename = image_map.get(rel_id)
            if not img_filename:
                continue
            rel_path = f"{source_stem}_images/{img_filename}"
            if parts and not parts[-1].endswith((" ", "\n")):
                parts.append(" ")
            parts.append(f"![{img_filename}]({rel_path})")
            parts.append(" ")

    return "".join(parts)


def _escape_markdown_cell(cell: str | None) -> str:
    """Normalize table cell content so it renders safely in Markdown tables."""
    if cell is None:
        return ""

    text = str(cell).replace("\r\n", "\n").replace("\r", "\n").strip()
    text = text.replace("\\", "\\\\")
    text = text.replace("\n", "<br>")
    text = text.replace("|", "\\|")
    return text


def _collect_files(root: Path, recursive: bool = False) -> list[Path]:
    """Collect convertible files from *root*, optionally recursing."""
    if recursive:
        files = [
            p for p in root.rglob("*")
            if p.suffix.lower() in SUPPORTED_EXTENSIONS and p.is_file()
        ]
    else:
        files = [
            p for p in root.iterdir()
            if p.suffix.lower() in SUPPORTED_EXTENSIONS and p.is_file()
        ]
    return sorted(files)


def _build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-docx-to-markdown",
        description="Convert PDF and DOCX files to clean Markdown.",
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Path to a single .pdf or .docx file to convert. "
             "If omitted, all supported files in the script directory are converted.",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default=None,
        help="Directory for the generated .md files (default: md_output/).",
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="In batch mode, recurse into subdirectories.",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable debug-level logging.",
    )
    parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress all output except errors.",
    )
    parser.add_argument(
        "--h1-threshold",
        type=int,
        default=6,
        metavar="PT",
        help="Font-size difference (in pt) above body text to classify as H1 (default: 6).",
    )
    parser.add_argument(
        "--h2-threshold",
        type=int,
        default=3,
        metavar="PT",
        help="Font-size difference (in pt) above body text to classify as H2 (default: 3).",
    )
    parser.add_argument(
        "--h3-threshold",
        type=int,
        default=1,
        metavar="PT",
        help="Font-size difference (in pt) above body text to classify as H3 (default: 1).",
    )
    parser.add_argument(
        "--font-sample-pages",
        type=int,
        default=20,
        metavar="N",
        help="Max pages sampled for body font-size detection (default: 20).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_argument_parser()
    args = parser.parse_args(argv)

    # Configure logging level.
    if args.quiet:
        log_level = logging.ERROR
    elif args.verbose:
        log_level = logging.DEBUG
    else:
        log_level = logging.INFO

    logging.basicConfig(level=log_level, format="%(message)s")

    cfg = ConversionConfig(
        heading1_threshold=args.h1_threshold,
        heading2_threshold=args.h2_threshold,
        heading3_threshold=args.h3_threshold,
        font_sample_pages=args.font_sample_pages,
    )

    output_dir = Path(args.output_dir) if args.output_dir else OUTPUT_DIR
    batch_start = time.time()

    # ── Single-file mode ──────────────────────────────────────────────
    if args.file:
        result = convert_document_to_markdown(args.file, output_dir, config=cfg)
        return 0 if result else 1

    # ── Batch mode ────────────────────────────────────────────────────
    files = _collect_files(SCRIPT_DIR, recursive=args.recursive)
    if not files:
        logger.info("No .docx or .pdf files found in %s", SCRIPT_DIR)
        return 0

    logger.info("\nFound %d file(s) to convert.", len(files))
    ok, failed = 0, 0

    for i, file in enumerate(files):
        result = convert_document_to_markdown(file, output_dir, config=cfg)
        if result:
            ok += 1
        else:
            failed += 1

        if i < len(files) - 1:
            remaining = len(files) - i - 1
            print(f"  Next: {files[i + 1].name}  ({remaining} file(s) remaining)")
            answer = input("  Continue? [y/n]: ").strip().lower()
            if answer != "y":
                print("\n  Stopped by user.")
                break

    total_elapsed = time.time() - batch_start
    mins, secs = divmod(int(total_elapsed), 60)
    time_str = f"{mins}m {secs}s" if mins else f"{secs:.1f}s"
    logger.info(
        "\n%s\n  ALL DONE - %d converted, %d failed\n  Total time: %s\n%s\n",
        "#" * 60, ok, failed, time_str, "#" * 60,
    )
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except FileNotFoundError as exc:
        logger.error("ERROR: %s", exc)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nStopped by user.")
        sys.exit(130)
    except Exception as exc:
        logger.error("ERROR: Unexpected failure: %s", exc)
        sys.exit(1)
