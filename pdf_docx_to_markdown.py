from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF - for image extraction from PDFs
import pdfplumber  # For text + table extraction from PDFs
from docx import Document  # python-docx - for DOCX conversion
from docx.oxml.ns import qn
from tqdm import tqdm

__version__ = "1.1.0"

SUPPORTED_EXTENSIONS = {".docx", ".pdf"}
SKIP_DIR_NAMES = {
    ".git",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    ".venv",
    "__pycache__",
    "md_output",
    "venv",
}
MIN_IMAGE_EDGE_PX = 20
MIN_IMAGE_BYTES = 400
_KIND_ORDER = {"text": 0, "table": 1, "image": 2}

_CONTENT_TYPE_EXT = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
    "image/tif": "tif",
    "image/webp": "webp",
    "image/svg+xml": "svg",
}
_SKIP_CONTENT_TYPES = {
    "image/x-emf",
    "image/x-wmf",
    "image/emf",
    "image/wmf",
    "application/x-msmetafile",
}

logger = logging.getLogger(__name__)


@dataclass
class ConversionConfig:
    """Configurable thresholds and limits for document conversion."""

    heading1_threshold: int = 6
    heading2_threshold: int = 3
    heading3_threshold: int = 2
    font_sample_pages: int = 20
    output_dir: Path | None = None


_DEFAULT_CONFIG = ConversionConfig()


def _default_output_dir() -> Path:
    return Path.cwd() / "md_output"


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
            source.suffix,
            ", ".join(sorted(SUPPORTED_EXTENSIONS)),
        )
        return None

    if output_dir:
        effective_output = Path(output_dir).expanduser().resolve()
    elif cfg.output_dir:
        effective_output = cfg.output_dir.expanduser().resolve()
    else:
        effective_output = _default_output_dir()
    effective_output.mkdir(parents=True, exist_ok=True)

    file_size_mb = source.stat().st_size / (1024 * 1024)
    logger.info(
        "\n%s\n  File    : %s\n  Size    : %.2f MB\n  Started : %s\n%s",
        "=" * 60,
        source.name,
        file_size_mb,
        datetime.now().strftime("%H:%M:%S"),
        "=" * 60,
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
        "=" * 60,
        md_path.name,
        out_size_kb,
        time_str,
        datetime.now().strftime("%H:%M:%S"),
        "=" * 60,
    )
    return str(md_path)


def _yaml_frontmatter(source: Path, extra: dict[str, Any] | None = None) -> str:
    lines = ["---", f"source: {source.name}"]
    if extra:
        for key, value in extra.items():
            lines.append(f"{key}: {value}")
    lines.append(f"extracted: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def _image_extension(content_type: str) -> str | None:
    """Map a DOCX image content type to a file extension, or None to skip."""
    ct = (content_type or "").lower().strip()
    if ct in _SKIP_CONTENT_TYPES:
        return None
    if ct in _CONTENT_TYPE_EXT:
        return _CONTENT_TYPE_EXT[ct]
    subtype = ct.split("/")[-1]
    if subtype in {"x-emf", "x-wmf", "emf", "wmf"}:
        return None
    if subtype in {"png", "jpeg", "jpg", "gif", "bmp", "tiff", "tif", "webp"}:
        return "jpg" if subtype == "jpeg" else subtype
    return None


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
            "Heading 6": "###### ",
        }

        images_dir = target_dir / f"{source.stem}_images"
        image_map: dict[str, str] = {}
        try:
            for rel in doc.part.rels.values():
                if "image" not in rel.reltype:
                    continue
                ext = _image_extension(rel.target_part.content_type)
                if ext is None:
                    logger.debug(
                        "Skipping unsupported DOCX image type %s",
                        rel.target_part.content_type,
                    )
                    continue
                img_data = rel.target_part.blob
                img_filename = f"img_{len(image_map) + 1}.{ext}"
                images_dir.mkdir(parents=True, exist_ok=True)
                (images_dir / img_filename).write_bytes(img_data)
                image_map[rel.rId] = img_filename
        except Exception as exc:
            logger.warning("DOCX image extraction failed: %s", exc)

        list_counters: dict[tuple[str, int], int] = {}
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
                content = _extract_docx_paragraph_content(
                    para, source.stem, image_map
                ).strip()

                if not content:
                    md_lines.append("")
                    continue

                style_name = para.style.name if para.style else ""
                prefix = heading_map.get(style_name, "")
                if prefix:
                    list_counters.clear()
                    md_lines.append(f"\n{prefix}{content}\n")
                    continue

                num_pr = _paragraph_num_pr(para, element)
                if num_pr is not None:
                    marker = _docx_list_marker(doc, num_pr, list_counters)
                    md_lines.append(f"{marker}{content}")
                elif style_name.startswith("List Number"):
                    key = (style_name, 0)
                    list_counters[key] = list_counters.get(key, 0) + 1
                    md_lines.append(f"{list_counters[key]}. {content}")
                elif style_name.startswith("List Bullet"):
                    md_lines.append(f"- {content}")
                else:
                    list_counters.clear()
                    md_lines.append(content)

            elif tag == "tbl":
                list_counters.clear()
                table = Table(element, doc)
                md_lines.append("")
                md_lines.extend(_docx_table_to_markdown(table, doc))
                md_lines.append("")

        if image_map:
            logger.info("  Extracted %d image(s) to %s", len(image_map), images_dir)
        elif images_dir.exists():
            try:
                images_dir.rmdir()
            except OSError:
                pass

        body = "\n".join(md_lines)
        return _yaml_frontmatter(source) + body
    except Exception as e:
        logger.error("python-docx conversion failed: %s", e)
        return None


def _paragraph_num_pr(para: Any, element: Any) -> Any | None:
    num_pr = element.find(".//" + qn("w:numPr"))
    if num_pr is not None:
        return num_pr
    style = para.style
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        try:
            num_pr = style.element.find(".//" + qn("w:numPr"))
        except Exception:
            num_pr = None
        if num_pr is not None:
            return num_pr
        try:
            style = style.base_style
        except Exception:
            break
    return None


def _docx_num_fmt(doc: Any, num_id: str, level: int) -> str:
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        return "bullet"

    abstract_id = None
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == num_id:
            abs_el = num.find(qn("w:abstractNumId"))
            if abs_el is not None:
                abstract_id = abs_el.get(qn("w:val"))
            break
    if abstract_id is None:
        return "bullet"

    for absn in numbering.findall(qn("w:abstractNum")):
        if absn.get(qn("w:abstractNumId")) != abstract_id:
            continue
        for lvl in absn.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) != str(level):
                continue
            fmt = lvl.find(qn("w:numFmt"))
            if fmt is not None:
                return str(fmt.get(qn("w:val"), "bullet"))
    return "bullet"


def _docx_list_marker(
    doc: Any,
    num_pr: Any,
    counters: dict[tuple[str, int], int],
) -> str:
    ilvl_el = num_pr.find(qn("w:ilvl"))
    num_id_el = num_pr.find(qn("w:numId"))
    level = int(ilvl_el.get(qn("w:val"), 0)) if ilvl_el is not None else 0
    num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else "0"
    fmt = _docx_num_fmt(doc, num_id, level)
    indent = "  " * level
    numbered = {
        "decimal",
        "decimalZero",
        "upperRoman",
        "lowerRoman",
        "upperLetter",
        "lowerLetter",
        "ordinal",
    }
    if fmt not in numbered:
        return f"{indent}- "
    key = (num_id, level)
    for other_level in [lvl for (_, lvl) in list(counters) if lvl > level]:
        counters.pop((num_id, other_level), None)
    counters[key] = counters.get(key, 0) + 1
    n = counters[key]
    if fmt in {"decimal", "decimalZero", "ordinal"}:
        return f"{indent}{n}. "
    if fmt == "upperLetter":
        return f"{indent}{chr(ord('A') + (n - 1) % 26)}. "
    if fmt == "lowerLetter":
        return f"{indent}{chr(ord('a') + (n - 1) % 26)}. "
    if fmt == "upperRoman":
        return f"{indent}{_to_roman(n)}. "
    if fmt == "lowerRoman":
        return f"{indent}{_to_roman(n).lower()}. "
    return f"{indent}{n}. "


def _to_roman(n: int) -> str:
    values = (
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    )
    n = max(1, min(n, 3999))
    out = []
    for value, numeral in values:
        while n >= value:
            out.append(numeral)
            n -= value
    return "".join(out)


def _format_docx_hyperlink(el: Any, para: Any) -> str:
    text = "".join((t.text or "") for t in el.findall(".//" + qn("w:t")))
    rel_id = el.get(qn("r:id"))
    url = ""
    if rel_id:
        try:
            url = para.part.rels[rel_id].target_ref
        except (KeyError, AttributeError):
            url = ""
    if url and text:
        return f"[{text}]({url})"
    return text


def _format_docx_run(
    run_el: Any, source_stem: str, image_map: dict[str, str], para: Any
) -> str:
    from docx.text.paragraph import Run

    run = Run(run_el, para)
    bits: list[str] = []
    if run.text:
        bits.append(run.text)
    for blip in run_el.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    ):
        rel_id = blip.get(qn("r:embed"))
        img_filename = image_map.get(rel_id) if rel_id else None
        if not img_filename:
            continue
        rel_path = f"{source_stem}_images/{img_filename}"
        if bits and not bits[-1].endswith((" ", "\n")):
            bits.append(" ")
        bits.append(f"![{img_filename}]({rel_path})")
        bits.append(" ")
    return "".join(bits)


def _extract_docx_paragraph_content(
    para: Any, source_stem: str, image_map: dict[str, str]
) -> str:
    """Return paragraph content with images and hyperlinks in document order."""
    parts: list[str] = []
    for child in para._element:
        tag = child.tag.split("}")[-1]
        if tag == "hyperlink":
            parts.append(_format_docx_hyperlink(child, para))
        elif tag == "r":
            parts.append(_format_docx_run(child, source_stem, image_map, para))
    if not parts:
        for run in para.runs:
            if run.text:
                parts.append(run.text)
    return "".join(parts)


def _should_keep_image(base_image: dict[str, Any], width: int = 0, height: int = 0) -> bool:
    data = base_image.get("image") or b""
    if len(data) < MIN_IMAGE_BYTES:
        return False
    img_w = int(base_image.get("width") or width or 0)
    img_h = int(base_image.get("height") or height or 0)
    if img_w and img_h and (img_w < MIN_IMAGE_EDGE_PX or img_h < MIN_IMAGE_EDGE_PX):
        return False
    return True


def _unique_image_xrefs(image_list: list[Any]) -> list[Any]:
    seen: set[int] = set()
    unique: list[Any] = []
    for img_info in image_list:
        xref = img_info[0]
        if xref in seen:
            continue
        seen.add(xref)
        unique.append(img_info)
    return unique


def _extract_pdf_page_images(
    fitz_doc: Any,
    page: Any,
    page_num: int,
    images_dir: Path,
) -> tuple[list[tuple[float, str]], int]:
    """Return (y, filename) blocks and the count of images written."""
    page_imgs: list[tuple[float, str]] = []
    written = 0
    image_list = _unique_image_xrefs(page.get_images(full=True))
    for img_idx, img_info in enumerate(image_list):
        xref = img_info[0]
        width = int(img_info[2]) if len(img_info) > 2 else 0
        height = int(img_info[3]) if len(img_info) > 3 else 0
        if width and height and (width < MIN_IMAGE_EDGE_PX or height < MIN_IMAGE_EDGE_PX):
            logger.debug(
                "Skipping tiny PDF image xref=%s (%sx%s) on page %s",
                xref,
                width,
                height,
                page_num + 1,
            )
            continue
        try:
            base_image = fitz_doc.extract_image(xref)
        except Exception as exc:
            logger.warning(
                "Failed to extract image xref=%s on page %s: %s",
                xref,
                page_num + 1,
                exc,
            )
            continue
        if not base_image or not base_image.get("image"):
            continue
        if not _should_keep_image(base_image, width, height):
            logger.debug("Skipping decorative PDF image xref=%s on page %s", xref, page_num + 1)
            continue
        ext = base_image.get("ext", "png")
        img_filename = f"page{page_num + 1}_img{img_idx + 1}.{ext}"
        images_dir.mkdir(parents=True, exist_ok=True)
        (images_dir / img_filename).write_bytes(base_image["image"])
        y = 0.0
        try:
            rects = page.get_image_rects(xref)
            if rects:
                y = float(rects[0].y0)
        except Exception as exc:
            logger.debug("Could not get image rect xref=%s: %s", xref, exc)
            y = float(page.rect.y1) if getattr(page, "rect", None) else 0.0
        page_imgs.append((y, img_filename))
        written += 1
    return page_imgs, written


def _convert_pdf(
    source: Path, target_dir: Path, cfg: ConversionConfig
) -> str | None:
    """Extract text, tables, and images from PDF using pdfplumber + PyMuPDF."""
    try:
        images_dir = target_dir / f"{source.stem}_images"
        page_image_blocks: dict[int, list[tuple[float, str]]] = {}
        image_count = 0

        with fitz.open(str(source)) as fitz_doc:
            total_pages = len(fitz_doc)
            for page_num in tqdm(
                range(total_pages),
                desc="  [1/2] Extracting images",
                unit="pg",
                ncols=60,
                ascii=True,
            ):
                page = fitz_doc[page_num]
                image_blocks, written = _extract_pdf_page_images(
                    fitz_doc, page, page_num, images_dir
                )
                page_image_blocks[page_num] = image_blocks
                image_count += written

        scanned_page_count = 0
        pages_blocks: list[list[tuple[float, str, str]]] = []

        with pdfplumber.open(str(source)) as pdf:
            body_font_size = _detect_body_font_size_from_pdf(pdf, cfg.font_sample_pages)
            for page_num, page in tqdm(
                enumerate(pdf.pages),
                desc="  [2/2] Extracting text ",
                unit="pg",
                ncols=60,
                total=total_pages,
                ascii=True,
            ):
                page_blocks = _extract_page_blocks(
                    page,
                    body_font_size,
                    cfg,
                    page_image_blocks.get(page_num, []),
                    source.stem,
                )
                pages_blocks.append(page_blocks)

        repeating = _detect_repeating_lines(
            [[c for _, kind, c in blocks if kind == "text"] for blocks in pages_blocks]
        )

        md_parts: list[str] = []
        for page_num, page_blocks in enumerate(pages_blocks):
            filtered: list[tuple[float, str, str]] = []
            has_text = False
            has_table = False
            has_image = False
            for y, kind, content in page_blocks:
                if kind == "text":
                    if _plain_line(content) in repeating:
                        continue
                    has_text = True
                elif kind == "table":
                    has_table = True
                elif kind == "image":
                    has_image = True
                filtered.append((y, kind, content))

            if has_image and not has_text and not has_table:
                scanned_page_count += 1

            body = _blocks_to_markdown(filtered)
            page_comment = f"<!-- Page {page_num + 1} -->"
            if not body.strip():
                continue
            md_parts.append(f"{page_comment}\n\n{body}".rstrip())

        if image_count > 0:
            logger.info("  Extracted %d image(s) to %s", image_count, images_dir)
        else:
            try:
                images_dir.rmdir()
            except OSError:
                pass

        if scanned_page_count > 0:
            pct = scanned_page_count / max(total_pages, 1) * 100
            logger.warning(
                "  WARNING: %d of %d page(s) (%.0f%%) appear to be scanned "
                "(images found but no extractable text).",
                scanned_page_count,
                total_pages,
                pct,
            )
            if scanned_page_count == total_pages:
                logger.warning(
                    "  This PDF seems to be fully scanned / image-based. "
                    "Consider using an OCR converter such as "
                    "https://github.com/Budzogan/pdf-image-ocr-to-markdown"
                )

        if not md_parts:
            logger.warning("No content extracted from %s", source.name)
            return _yaml_frontmatter(source, {"pages": total_pages})

        metadata = _yaml_frontmatter(source, {"pages": total_pages})
        return metadata + "\n\n---\n\n".join(md_parts)
    except Exception as e:
        logger.error("PDF extraction failed: %s", e)
        return None


def _detect_body_font_size_from_pdf(pdf: Any, sample_pages: int = 20) -> int:
    all_sizes: list[int] = []
    try:
        for page in pdf.pages[:sample_pages]:
            try:
                words = page.extract_words(extra_attrs=["size"])
            except Exception:
                continue
            all_sizes.extend(round(word["size"]) for word in words if word.get("size"))
    except Exception:
        return 10
    if not all_sizes:
        return 10
    return Counter(all_sizes).most_common(1)[0][0]


def _detect_body_font_size(source: Path, sample_pages: int = 20) -> int:
    """Scan the PDF and return the most common font size (= body text)."""
    try:
        with pdfplumber.open(str(source)) as pdf:
            return _detect_body_font_size_from_pdf(pdf, sample_pages)
    except Exception:
        return 10


def _plain_line(text: str) -> str:
    return re.sub(r"^#{1,6}\s*", "", text).strip()


def _detect_repeating_lines(pages_lines: list[list[str]]) -> set[str]:
    """Return header/footer lines that repeat across many pages."""
    n = len(pages_lines)
    if n < 3:
        return set()
    candidates: Counter[str] = Counter()
    for lines in pages_lines:
        plains = [_plain_line(line) for line in lines if _plain_line(line)]
        edge: list[str] = []
        if plains:
            edge.extend(plains[:2])
            if len(plains) > 2:
                edge.extend(plains[-2:])
        for text in dict.fromkeys(edge):
            if len(text) >= 4:
                candidates[text] += 1
    threshold = max(3, int(n * 0.5))
    return {text for text, count in candidates.items() if count >= threshold}


def _extract_text_lines_with_headings(
    page: Any, body_font_size: int, cfg: ConversionConfig | None = None
) -> list[tuple[float, str]]:
    if cfg is None:
        cfg = _DEFAULT_CONFIG

    try:
        words = page.extract_words(extra_attrs=["size"])
    except Exception:
        text = page.extract_text() or ""
        return [(0.0, text)] if text else []

    if not words:
        return []

    line_buckets: dict[int, list[dict[str, Any]]] = {}
    for word in words:
        top = round(word["top"] / 3) * 3
        line_buckets.setdefault(top, []).append(word)

    result_lines: list[tuple[float, str]] = []
    for top in sorted(line_buckets):
        line_words = sorted(line_buckets[top], key=lambda item: item.get("x0", 0))
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

        result_lines.append((float(top), text))

    return result_lines


def _extract_text_with_headings(
    page: Any, body_font_size: int, cfg: ConversionConfig | None = None
) -> str:
    """Extract page text, promoting heading lines to markdown # notation."""
    lines = _extract_text_lines_with_headings(page, body_font_size, cfg)
    return "\n".join(text for _, text in lines)


def _extract_page_blocks(
    page: Any,
    body_font_size: int,
    cfg: ConversionConfig,
    image_blocks: list[tuple[float, str]],
    source_stem: str,
) -> list[tuple[float, str, str]]:
    blocks: list[tuple[float, str, str]] = []
    table_bboxes: list[tuple[float, float, float, float]] = []

    try:
        found_tables = page.find_tables()
    except Exception as exc:
        logger.debug("Table detection failed: %s", exc)
        found_tables = []

    for table in found_tables:
        try:
            bbox = table.bbox
            table_bboxes.append(bbox)
            data = table.extract()
        except Exception as exc:
            logger.debug("Table extract failed: %s", exc)
            continue
        if data:
            md = _table_to_markdown(data)
            if md:
                blocks.append((float(bbox[1]), "table", md))

    filtered_page = page
    for bbox in table_bboxes:
        try:
            filtered_page = filtered_page.outside_bbox(bbox)
        except Exception as exc:
            logger.debug("outside_bbox failed: %s", exc)

    for y, text in _extract_text_lines_with_headings(filtered_page, body_font_size, cfg):
        if text.strip():
            blocks.append((y, "text", text))

    for y, filename in image_blocks:
        rel_path = f"{source_stem}_images/{filename}"
        blocks.append((y, "image", f"![Image]({rel_path})"))

    return blocks


def _blocks_to_markdown(blocks: list[tuple[float, str, str]]) -> str:
    if not blocks:
        return ""
    ordered = sorted(blocks, key=lambda item: (item[0], _KIND_ORDER.get(item[1], 9)))
    lines: list[str] = []
    prev_kind: str | None = None
    for _, kind, content in ordered:
        if not str(content).strip():
            continue
        if prev_kind is not None and not (prev_kind == "text" and kind == "text"):
            lines.append("")
        lines.append(content)
        prev_kind = kind
    return "\n".join(lines)


def _table_to_markdown(table: list[list[str | None]] | None) -> str:
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

    for clean_row in clean[1:]:
        while len(clean_row) < len(clean[0]):
            clean_row.append("")
        lines.append("| " + " | ".join(clean_row[: len(clean[0])]) + " |")

    return "\n".join(lines)


def _docx_table_to_markdown(table: Any, doc: Any) -> list[str]:
    """Convert a python-docx Table to markdown lines, handling merged cells and nested tables."""
    from docx.table import Table as DocxTable

    rows = table.rows
    if not rows:
        return []

    def _dedup_row(row: Any) -> list[str]:
        seen_ids: set[int] = set()
        cells: list[str] = []
        for cell in row.cells:
            cid = id(cell._element)
            if cid in seen_ids:
                cells.append("")
            else:
                seen_ids.add(cid)
                nested = cell._element.findall(qn("w:tbl"))
                if nested:
                    parts = [cell.text.strip()]
                    for ntbl_el in nested:
                        ntbl = DocxTable(ntbl_el, doc)
                        nested_lines = _docx_table_to_markdown(ntbl, doc)
                        parts.append(" ".join(nested_lines))
                    cells.append(_escape_markdown_cell(" | ".join(p for p in parts if p)))
                else:
                    cells.append(_escape_markdown_cell(cell.text))
        return cells

    header = _dedup_row(rows[0])
    col_count = len(header)
    lines: list[str] = []
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * col_count) + " |")

    for row in rows[1:]:
        cells = _dedup_row(row)
        while len(cells) < col_count:
            cells.append("")
        cells = cells[:col_count]
        lines.append("| " + " | ".join(cells) + " |")

    return lines


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
        files: list[Path] = []
        for path in root.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            try:
                relative_parts = path.relative_to(root).parts[:-1]
            except ValueError:
                relative_parts = ()
            if any(part in SKIP_DIR_NAMES or part.startswith(".") for part in relative_parts):
                continue
            files.append(path)
        return sorted(files)

    return sorted(
        path
        for path in root.iterdir()
        if path.suffix.lower() in SUPPORTED_EXTENSIONS and path.is_file()
    )


def _should_prompt_between_files(assume_yes: bool) -> bool:
    if assume_yes:
        return False
    isatty = getattr(sys.stdin, "isatty", None)
    return bool(isatty and isatty())


def _prompt_continue(next_name: str, remaining: int) -> bool:
    print(f"  Next: {next_name}  ({remaining} file(s) remaining)")
    answer = input("  Continue? [Y/n]: ").strip().lower()
    return answer in {"", "y", "yes"}


def _build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-docx-to-markdown",
        description="Convert PDF and DOCX files to clean Markdown.",
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Path to a single .pdf or .docx file to convert. "
        "If omitted, all supported files in the current directory are converted.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help="Directory for the generated .md files (default: ./md_output/).",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="In batch mode, recurse into subdirectories.",
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="List files that would be converted without actually converting.",
    )
    parser.add_argument(
        "-y",
        "--yes",
        action="store_true",
        help="Do not prompt between files in batch mode.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable debug-level logging.",
    )
    parser.add_argument(
        "-q",
        "--quiet",
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
        default=2,
        metavar="PT",
        help="Font-size difference (in pt) above body text to classify as H3 (default: 2).",
    )
    parser.add_argument(
        "--font-sample-pages",
        type=int,
        default=20,
        metavar="N",
        help="Max pages sampled for body font-size detection (default: 20).",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_argument_parser()
    args = parser.parse_args(argv)

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

    output_dir = Path(args.output_dir) if args.output_dir else _default_output_dir()
    batch_start = time.time()
    batch_root = Path.cwd()

    if args.file:
        if args.dry_run:
            source = Path(args.file).expanduser().resolve()
            if not source.exists():
                logger.error("File not found: %s", source)
                return 1
            if source.suffix.lower() not in {".pdf", ".docx"}:
                logger.error(
                    "Unsupported file type: %s (expected .pdf or .docx)", source.name
                )
                return 1
            size_mb = source.stat().st_size / (1024 * 1024)
            logger.info("  [dry-run] Would convert: %s (%.2f MB)", source.name, size_mb)
            return 0
        result = convert_document_to_markdown(args.file, output_dir, config=cfg)
        return 0 if result else 1

    files = _collect_files(batch_root, recursive=args.recursive)
    if not files:
        logger.info("No .docx or .pdf files found in %s", batch_root)
        return 0

    if args.dry_run:
        logger.info("\n[dry-run] %d file(s) would be converted:\n", len(files))
        total_size = 0.0
        for path in files:
            size_mb = path.stat().st_size / (1024 * 1024)
            total_size += size_mb
            rel = path.relative_to(batch_root) if path.is_relative_to(batch_root) else path
            logger.info("  %-50s  %.2f MB", rel, size_mb)
        logger.info("\n  Total: %.2f MB across %d file(s)", total_size, len(files))
        return 0

    logger.info("\nFound %d file(s) to convert.", len(files))
    ok, failed = 0, 0
    prompt = _should_prompt_between_files(args.yes)

    for i, file in enumerate(files):
        result = convert_document_to_markdown(file, output_dir, config=cfg)
        if result:
            ok += 1
        else:
            failed += 1

        if i < len(files) - 1 and prompt:
            remaining = len(files) - i - 1
            if not _prompt_continue(files[i + 1].name, remaining):
                print("\n  Stopped by user.")
                break

    total_elapsed = time.time() - batch_start
    mins, secs = divmod(int(total_elapsed), 60)
    time_str = f"{mins}m {secs}s" if mins else f"{secs:.1f}s"
    logger.info(
        "\n%s\n  ALL DONE - %d converted, %d failed\n  Total time: %s\n%s\n",
        "#" * 60,
        ok,
        failed,
        time_str,
        "#" * 60,
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
