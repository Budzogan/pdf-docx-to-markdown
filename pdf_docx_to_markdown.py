import logging
import re
import sys
import time
from collections import Counter
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


def convert_document_to_markdown(
    input_path: str | Path, output_dir: str | Path | None = None
) -> str | None:
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"File not found: {source}")

    target_dir = Path(output_dir).expanduser().resolve() if output_dir else source.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    file_size_mb = source.stat().st_size / (1024 * 1024)
    logger.info(
        "\n%s\n  File    : %s\n  Size    : %.2f MB\n  Started : %s\n%s",
        "=" * 60, source.name, file_size_mb,
        datetime.now().strftime("%H:%M:%S"), "=" * 60,
    )

    t_start = time.time()

    if source.suffix.lower() == ".pdf":
        markdown_content = _convert_pdf(source, target_dir)
    else:
        logger.info("  Processing DOCX...")
        markdown_content = _convert_with_python_docx(source, target_dir)

    elapsed = time.time() - t_start

    if markdown_content is None:
        logger.error("  ERROR: Failed to convert %s.", source.name)
        return None

    md_path = target_dir / f"{source.stem}.md"
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

        for element in doc.element.body:
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


def _convert_pdf(source: Path, target_dir: Path) -> str | None:
    """Extract text, tables, and images from PDF using pdfplumber + PyMuPDF."""
    try:
        images_dir = target_dir / f"{source.stem}_images"
        images_dir.mkdir(parents=True, exist_ok=True)

        md_parts: list[str] = []
        image_count = 0
        body_font_size = _detect_body_font_size(source)

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

                text = _extract_text_with_headings(filtered_page, body_font_size)
                if text.strip():
                    page_md += text.strip() + "\n\n"

                for table in tables:
                    if table and len(table) > 0:
                        page_md += _table_to_markdown(table) + "\n\n"

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


def _detect_body_font_size(source: Path) -> int:
    """Scan the whole PDF and return the most common font size (= body text)."""
    try:
        all_sizes: list[int] = []
        with pdfplumber.open(str(source)) as pdf:
            for page in pdf.pages[:20]:
                words = page.extract_words(extra_attrs=["size"])
                all_sizes.extend(round(word["size"]) for word in words if word.get("size"))
        if not all_sizes:
            return 10
        return Counter(all_sizes).most_common(1)[0][0]
    except Exception:
        return 10


def _extract_text_with_headings(page: object, body_font_size: int) -> str:
    """Extract page text, promoting heading lines to markdown # notation."""
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

        if diff >= 6:
            text = f"# {text}"
        elif diff >= 3:
            text = f"## {text}"
        elif diff >= 1:
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


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
    )

    batch_start = time.time()

    if len(sys.argv) > 1:
        input_path = sys.argv[1]
        output_dir = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_DIR
        result = convert_document_to_markdown(input_path, output_dir)
        return 0 if result else 1

    files = [file for file in SCRIPT_DIR.iterdir() if file.suffix.lower() in SUPPORTED_EXTENSIONS]
    if not files:
        logger.info("No .docx or .pdf files found in %s", SCRIPT_DIR)
        return 0

    logger.info("\nFound %d file(s) to convert.", len(files))
    ok, failed = 0, 0

    for i, file in enumerate(files):
        result = convert_document_to_markdown(file, OUTPUT_DIR)
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
