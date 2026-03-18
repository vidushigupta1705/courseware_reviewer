import os
import re
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from models import ParagraphInfo, ImageInfo, DocumentState, TableInfo, TableCellInfo
from utils import pt_to_float, normalize_text, has_extra_spacing


ALIGNMENT_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
    WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
    WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
}


def detect_heading_level(style_name: str) -> Optional[int]:
    if not style_name:
        return None
    style_name = style_name.strip().lower()
    if style_name.startswith("heading"):
        parts = style_name.split()
        if len(parts) >= 2 and parts[1].isdigit():
            level = int(parts[1])
            if 1 <= level <= 9:
                return level
    return None


def is_code_like_paragraph(text: str, style_name: str = "") -> bool:
    if not text or not text.strip():
        return False

    lower_style = (style_name or "").lower()
    if any(x in lower_style for x in ["code", "source code", "preformatted"]):
        return True

    score = 0
    if re.search(r"[{}();=<>\[\]#]", text):
        score += 1
    if re.search(r"\bdef\s+\w+\s*\(", text):
        score += 2
    if re.search(r"\bclass\s+\w+", text):
        score += 2
    if re.search(r"\bimport\s+\w+", text):
        score += 2
    if re.search(r"\bfor\b.*:", text):
        score += 1
    if re.search(r"\bif\b.*:", text):
        score += 1
    if re.search(r"^\s{2,}\S+", text, flags=re.MULTILINE):
        score += 1
    if re.search(r"<[^>]+>", text):
        score += 1

    return score >= 3


def extract_paragraphs(doc: Document) -> List[ParagraphInfo]:
    paragraphs = []

    for idx, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else "Normal"
        heading_level = detect_heading_level(style_name)
        is_heading = heading_level is not None

        font_names = []
        font_sizes = []
        bold_flags = []

        for run in para.runs:
            if run.text.strip():
                if run.font.name:
                    font_names.append(run.font.name)
                elif run.style and run.style.font and run.style.font.name:
                    font_names.append(run.style.font.name)

                if run.font.size:
                    size_val = pt_to_float(run.font.size)
                    if size_val is not None:
                        font_sizes.append(size_val)

                if run.bold is not None:
                    bold_flags.append(bool(run.bold))

        text = normalize_text(para.text)

        paragraph_info = ParagraphInfo(
            index=idx,
            text=text,
            style_name=style_name,
            is_heading=is_heading,
            heading_level=heading_level,
            font_names=font_names,
            font_sizes=font_sizes,
            bold_flags=bold_flags,
            alignment=ALIGNMENT_MAP.get(para.alignment, "UNKNOWN"),
            has_extra_spacing_issue=has_extra_spacing(para.text),
            is_code_like=is_code_like_paragraph(para.text, style_name),
        )
        paragraphs.append(paragraph_info)

    return paragraphs


def extract_tables(doc: Document) -> List[TableInfo]:
    tables = []

    # Walk body children to find the paragraph index just before each table.
    # Body children are a mix of <w:p> (paragraphs) and <w:tbl> (tables).
    W_P = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
    W_TBL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"

    para_count = 0
    table_para_positions = []  # paragraph index just before each table

    for child in doc.element.body:
        if child.tag == W_P:
            para_count += 1
        elif child.tag == W_TBL:
            # anchor to the last paragraph seen before this table (min 0)
            table_para_positions.append(max(0, para_count - 1))

    for t_idx, table in enumerate(doc.tables):
        row_count = len(table.rows)
        col_count = max((len(r.cells) for r in table.rows), default=0)

        cell_items = []
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                cell_items.append(
                    TableCellInfo(
                        row_index=r_idx,
                        col_index=c_idx,
                        text=normalize_text(cell.text),
                    )
                )

        para_idx = table_para_positions[t_idx] if t_idx < len(table_para_positions) else None

        tables.append(
            TableInfo(
                table_index=t_idx,
                row_count=row_count,
                col_count=col_count,
                cells=cell_items,
                paragraph_index=para_idx,
            )
        )

    return tables


def _get_rel_ids_from_paragraph(paragraph):
    rel_ids = []
    # Only look inside w:drawing elements — these are real embedded images.
    # w:pict elements contain bullets, themes, and background shapes — not content images.
    drawings = paragraph._element.xpath('.//*[local-name()="drawing"]')
    for drawing in drawings:
        blips = drawing.xpath('.//*[local-name()="blip"]')
        for blip in blips:
            for attr_name, attr_value in blip.attrib.items():
                if attr_name.endswith('}embed'):
                    rel_ids.append(attr_value)
    return rel_ids


def build_image_paragraph_map(doc: Document):
    """
    Returns dict: rel_id -> paragraph index

    Scans both regular paragraphs and table cells for embedded images.
    """
    rel_to_para = {}

    # Scan regular paragraphs
    for p_idx, para in enumerate(doc.paragraphs):
        rel_ids = _get_rel_ids_from_paragraph(para)
        if rel_ids:
            for rel_id in rel_ids:
                if rel_id not in rel_to_para:
                    rel_to_para[rel_id] = p_idx

    # Scan table cells — images inside tables won't appear in doc.paragraphs
    # Find the nearest paragraph index before each table using body element order
    W_P = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
    W_TBL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"

    para_count = 0
    for child in doc.element.body:
        if child.tag == W_P:
            para_count += 1
        elif child.tag == W_TBL:
            anchor_idx = max(0, para_count - 1)
            # Only find blips inside w:drawing elements — real embedded images
            drawings = child.xpath('.//*[local-name()="drawing"]')
            for drawing in drawings:
                blips = drawing.xpath('.//*[local-name()="blip"]')
                for blip in blips:
                    for attr_name, attr_value in blip.attrib.items():
                        if attr_name.endswith('}embed'):
                            rel_id = attr_value
                            if rel_id not in rel_to_para:
                                rel_to_para[rel_id] = anchor_idx

    return rel_to_para


def extract_images(doc: Document, output_dir: Path) -> List[ImageInfo]:
    output_dir.mkdir(parents=True, exist_ok=True)
    images = []
    rels = doc.part.rels
    image_para_map = build_image_paragraph_map(doc)

    seen = set()

    for rel_id, rel in rels.items():
        if "image" not in rel.target_ref:
            continue

        # Only process images that are actually embedded in the document body
        # (paragraphs or tables). Images not in image_para_map are theme images,
        # bullet images, backgrounds etc. — not content images.
        if rel_id not in image_para_map:
            continue

        image_part = rel.target_part
        image_bytes = image_part.blob

        # Skip tiny images — these are typically invisible 1x1 tracking pixels,
        # auto-inserted transparent images, or Word internal assets.
        # Real content images are always larger than 1KB.
        if len(image_bytes) < 1024:
            continue
        original_name = os.path.basename(rel.target_ref)
        if not original_name:
            original_name = f"{rel_id}.png"

        output_path = output_dir / original_name
        candidate = output_path
        count = 1
        while candidate.exists() or str(candidate) in seen:
            stem = output_path.stem
            suffix = output_path.suffix
            candidate = output_dir / f"{stem}_{count}{suffix}"
            count += 1

        with open(candidate, "wb") as f:
            f.write(image_bytes)

        seen.add(str(candidate))
        images.append(
            ImageInfo(
                rel_id=rel_id,
                filename=os.path.basename(candidate),
                output_path=str(candidate),
                paragraph_index=image_para_map.get(rel_id),
            )
        )

    return images

    return images


def load_docx_state(input_file: Path, extracted_images_dir: Path) -> DocumentState:
    doc = Document(str(input_file))
    base_filename = input_file.stem

    paragraphs = extract_paragraphs(doc)
    tables = extract_tables(doc)
    images = extract_images(doc, extracted_images_dir / base_filename)

    return DocumentState(
        input_file=str(input_file),
        base_filename=base_filename,
        paragraphs=paragraphs,
        tables=tables,
        images=images,
        ocr_results=[],
        rewrite_suggestions=[],
        units=[],
        issues=[],
    )