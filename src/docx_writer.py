from pathlib import Path
import re

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches

from config import (
    REQUIRED_HEADING_FONT,
    REQUIRED_HEADING_SIZE_PT,
    REQUIRED_HEADING_BOLD,
    REQUIRED_BODY_FONT,
    REQUIRED_BODY_SIZE_PT,
    COMMENT_AUTHOR,
    COMMENT_INITIALS,
    DIAGRAM_IMAGE_WIDTH_INCHES,
    ADVANCED_VISUAL_WIDTH_INCHES,
)
from utils import clean_spacing, normalize_text


def _get_dominant_font(doc: Document, is_heading: bool):
    """Find the most commonly used font name and size for headings or body paragraphs."""
    from collections import Counter
    font_names = []
    font_sizes = []
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        para_is_heading = style_name.lower().startswith("heading")
        if para_is_heading != is_heading:
            continue
        if not para.text.strip():
            continue
        for run in para.runs:
            if run.text.strip():
                if run.font.name:
                    font_names.append(run.font.name)
                if run.font.size:
                    font_sizes.append(run.font.size.pt)
    dominant_font = Counter(font_names).most_common(1)[0][0] if font_names else None
    dominant_size = Counter(font_sizes).most_common(1)[0][0] if font_sizes else None
    return dominant_font, dominant_size


def normalize_styles(doc: Document):
    """
    Fix only paragraphs whose font/size is inconsistent with the document's
    own dominant style. Does not force Arial on documents that consistently
    use a different font — only corrects genuine outliers.
    Also updates TOC style definitions to use the dominant body font.
    """
    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)

    # Update TOC style definitions so auto-generated or manually styled
    # TOC entries use the document's dominant font
    toc_style_names = ["toc 1", "toc 2", "toc 3", "toc 4", "toc 5",
                       "toc 6", "toc 7", "toc 8", "toc 9", "TOC Heading"]
    target_toc_font = dom_body_font or REQUIRED_BODY_FONT
    for style_name in toc_style_names:
        try:
            style = doc.styles[style_name]
            if style.font:
                style.font.name = target_toc_font
        except KeyError:
            pass

    # Also fix any paragraphs already using toc styles
    for para in doc.paragraphs:
        style_name = (para.style.name if para.style else "").lower()
        if style_name.startswith("toc"):
            for run in para.runs:
                if run.text.strip():
                    run.font.name = target_toc_font

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        is_heading = style_name.lower().startswith("heading")

        target_font = dom_heading_font if is_heading else dom_body_font
        target_size = dom_heading_size if is_heading else dom_body_size

        if not target_font and not target_size:
            continue

        for run in para.runs:
            if not run.text.strip():
                continue

            run_font = run.font.name
            run_size = run.font.size.pt if run.font.size else None

            font_is_outlier = (
                run_font and target_font and run_font != target_font
            )
            size_is_outlier = (
                run_size is not None and target_size is not None
                and abs(run_size - target_size) > 0.5
            )

            if font_is_outlier:
                run.font.name = target_font
            if size_is_outlier:
                run.font.size = Pt(target_size)
            if is_heading:
                run.bold = True


def clean_doc_spacing(doc: Document):
    for para in doc.paragraphs:
        # Never clear a paragraph that contains an embedded image
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue

        if para.text and para.text.strip():
            cleaned = clean_spacing(para.text)
            if cleaned != para.text.strip():
                para.clear()
                new_run = para.add_run(cleaned)
                style_name = para.style.name if para.style else "Normal"
                if style_name.lower().startswith("heading"):
                    new_run.font.name = REQUIRED_HEADING_FONT
                    new_run.font.size = Pt(REQUIRED_HEADING_SIZE_PT)
                    new_run.bold = REQUIRED_HEADING_BOLD
                else:
                    new_run.font.name = REQUIRED_BODY_FONT
                    new_run.font.size = Pt(REQUIRED_BODY_SIZE_PT)


def copy_docx(input_path: Path) -> Document:
    return Document(str(input_path))


def _paragraph_has_meaningful_text(paragraph):
    return bool(paragraph.text and paragraph.text.strip())


def _get_best_paragraph(doc: Document, paragraph_index):
    if not doc.paragraphs:
        p = doc.add_paragraph(" ")
        return p, 0

    if paragraph_index is not None and 0 <= paragraph_index < len(doc.paragraphs):
        p = doc.paragraphs[paragraph_index]
        if _paragraph_has_meaningful_text(p) or p.runs:
            return p, paragraph_index

        for distance in range(1, 6):
            prev_idx = paragraph_index - distance
            next_idx = paragraph_index + distance

            if prev_idx >= 0:
                prev_p = doc.paragraphs[prev_idx]
                if _paragraph_has_meaningful_text(prev_p) or prev_p.runs:
                    return prev_p, prev_idx

            if next_idx < len(doc.paragraphs):
                next_p = doc.paragraphs[next_idx]
                if _paragraph_has_meaningful_text(next_p) or next_p.runs:
                    return next_p, next_idx

    for idx, p in enumerate(doc.paragraphs):
        if _paragraph_has_meaningful_text(p) or p.runs:
            return p, idx

    p = doc.paragraphs[0]
    return p, 0


def _get_or_create_anchor_run(paragraph):
    for run in paragraph.runs:
        if run.text and run.text.strip():
            return run
    return paragraph.add_run(" ")


def _highlight_anchor_context(paragraph, max_runs=2):
    highlighted = 0
    for run in paragraph.runs:
        if run.text and run.text.strip():
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            highlighted += 1
            if highlighted >= max_runs:
                break

    if highlighted == 0 and paragraph.runs:
        paragraph.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW


def _safe_comment_text(issue):
    return (
        f"Issue Type: {issue.issue_type}\n"
        f"Severity: {issue.severity}\n"
        f"Location: {issue.location}\n"
        f"Comment: {issue.message}\n"
        f"Suggested Fix: {issue.suggested_fix}"
    )


def add_inline_comments_to_doc(doc: Document, state):
    if not doc.paragraphs:
        doc.add_paragraph(" ")

    for issue in state.issues:
        paragraph, _ = _get_best_paragraph(doc, issue.paragraph_index)
        anchor_run = _get_or_create_anchor_run(paragraph)
        _highlight_anchor_context(paragraph)

        try:
            doc.add_comment(
                anchor_run,
                text=_safe_comment_text(issue),
                author=COMMENT_AUTHOR,
                initials=COMMENT_INITIALS,
            )
        except Exception:
            fallback_run = paragraph.add_run(" ")
            fallback_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            doc.add_comment(
                fallback_run,
                text=_safe_comment_text(issue),
                author=COMMENT_AUTHOR,
                initials=COMMENT_INITIALS,
            )


def _is_figure_caption(text: str) -> bool:
    if not text:
        return False
    return bool(re.search(r'^\s*(figure|fig\.?|fig)\s*\d+', text, flags=re.IGNORECASE))


def _get_next_figure_number(doc: Document) -> int:
    """
    Scan the document for existing figure numbers and return
    the next sequential number to use.
    """
    max_fig = 0
    for para in doc.paragraphs:
        m = re.search(r'figure\s*(\d+)', para.text, flags=re.IGNORECASE)
        if m:
            num = int(m.group(1))
            if num > max_fig:
                max_fig = num
    return max_fig + 1


def _append_caption_then_source(doc: Document, idx: int, figure_text: str, source_text: str):
    """
    Insert figure caption then source in the correct order:
    Image → Caption (Figure N) → Source

    The caller is responsible for deduplication via source_inserted_at /
    caption_inserted_at sets. This function inserts unconditionally when
    figure_text / source_text are provided, but still skips if an existing
    caption is already present (to avoid duplicating real captions from the doc).
    """
    if idx is None or idx < 0 or idx >= len(doc.paragraphs):
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
    else:
        anchor = doc.paragraphs[idx]

    # Only check for existing caption — to avoid duplicating real figure labels
    def _has_caption_nearby(para_idx):
        for check_idx in range(para_idx, min(para_idx + 4, len(doc.paragraphs))):
            if _is_figure_caption(doc.paragraphs[check_idx].text):
                return True
        return False

    caption_anchor = anchor

    # Insert caption first (Image → Caption)
    if figure_text and not _has_caption_nearby(idx if idx is not None else 0):
        new_cap = insert_paragraph_after(caption_anchor)
        run = new_cap.add_run(figure_text)
        run.italic = True
        caption_anchor = new_cap

    # Insert source after caption (Caption → Source)
    # No duplicate check here — caller controls this via source_inserted_at
    if source_text:
        new_src = insert_paragraph_after(caption_anchor)
        run = new_src.add_run(source_text)
        run.italic = True


def add_ocr_summary_to_review_doc(doc: Document, state):
    # OCR and rewrite summaries are internal pipeline info.
    # Not shown in the Review Comments document.
    pass


def add_missing_summary_headings(doc: Document):
    """
    Detects paragraphs that begin with summary/conclusion phrases and
    inserts a heading before them if one is not already present.
    """
    import re

    SUMMARY_PATTERNS = [
        r'^in summary[,\s]',
        r'^to summarize[,\s]',
        r'^in conclusion[,\s]',
        r'^to conclude[,\s]',
        r'^summary[:\s]',
        r'^conclusion[:\s]',
        r'^overall[,\s]',
        r'^in brief[,\s]',
    ]

    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    h_font = dom_heading_font or REQUIRED_HEADING_FONT
    h_size = dom_heading_size or REQUIRED_HEADING_SIZE_PT

    # Collect indices first to avoid mutation during iteration
    to_insert = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip().lower()
        style = para.style.name if para.style else "Normal"

        if style.lower().startswith("heading"):
            continue

        if any(re.match(p, text) for p in SUMMARY_PATTERNS):
            # Check if the paragraph immediately before is a heading
            if i > 0:
                prev_style = doc.paragraphs[i - 1].style.name if doc.paragraphs[i - 1].style else ""
                if prev_style.lower().startswith("heading"):
                    continue
            to_insert.append(i)

    # Insert headings from bottom to top to preserve indices
    for i in sorted(to_insert, reverse=True):
        anchor = doc.paragraphs[i]
        # Insert "Summary" heading before this paragraph
        new_p = OxmlElement("w:p")
        anchor._p.addprevious(new_p)
        from docx.text.paragraph import Paragraph
        new_para = Paragraph(new_p, anchor._parent)
        new_para.style = doc.styles["Heading 3"]
        run = new_para.add_run("Summary")
        run.font.name = h_font
        run.font.size = Pt(h_size)
        run.bold = True
    # OCR and rewrite summaries are internal pipeline info.
    # Not shown in the Review Comments document.
    pass


def _remove_paragraph(paragraph):
    p = paragraph._element
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def remove_exact_duplicate_body_paragraphs(doc: Document):
    """
    Remove duplicate body paragraphs that appear within a close range of each other.
    Only removes duplicates within MAX_DUPLICATE_DISTANCE paragraphs — duplicates
    that are far apart are likely intentional repetition (e.g. same example used
    in two different sections) and should NOT be removed.
    """
    MAX_DUPLICATE_DISTANCE = 20

    seen = {}  # norm_text -> paragraph_index
    to_remove = []

    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else "Normal"
        is_heading = style_name.lower().startswith("heading")
        norm = normalize_text(para.text).lower()

        if not norm or is_heading:
            continue

        if len(norm) < 25:
            continue

        # Never remove paragraphs containing images
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue

        if norm in seen:
            first_idx = seen[norm]
            distance = i - first_idx
            if distance <= MAX_DUPLICATE_DISTANCE:
                to_remove.append(para)
            # Update to the most recent occurrence for future comparisons
            seen[norm] = i
        else:
            seen[norm] = i

    for para in to_remove:
        _remove_paragraph(para)


def apply_rewrites_to_doc(doc: Document, state):
    for item in getattr(state, "rewrite_suggestions", []):
        if item.skipped or not item.applied:
            continue
        idx = item.paragraph_index
        if idx is None or idx < 0 or idx >= len(doc.paragraphs):
            continue

        para = doc.paragraphs[idx]

        # Never rewrite a paragraph that contains an embedded image —
        # para.clear() would destroy the image
        has_image = para._element.xpath('.//*[local-name()="drawing"] | .//*[local-name()="pict"]')
        if has_image:
            continue

        style_name = para.style.name if para.style else "Normal"
        is_heading = style_name.lower().startswith("heading")

        para.clear()
        run = para.add_run(item.rewritten_text)

        if is_heading:
            run.font.name = REQUIRED_HEADING_FONT
            run.font.size = Pt(REQUIRED_HEADING_SIZE_PT)
            run.bold = REQUIRED_HEADING_BOLD
        else:
            run.font.name = REQUIRED_BODY_FONT
            run.font.size = Pt(REQUIRED_BODY_SIZE_PT)

def append_generated_quizzes(doc: Document, state):
    from docx.shared import Pt

    for unit in getattr(state, "units", []):
        if unit.quiz_present or not unit.generated_quiz:
            continue

        heading = doc.add_paragraph()
        r = heading.add_run(f"Quiz – {unit.title}")
        r.bold = True
        r.font.name = REQUIRED_HEADING_FONT
        r.font.size = Pt(REQUIRED_HEADING_SIZE_PT)

        current_type = None
        for item in unit.generated_quiz:
            if item.qtype != current_type:
                current_type = item.qtype
                sub = doc.add_paragraph()
                sr = sub.add_run(current_type)
                sr.bold = True
                sr.font.name = REQUIRED_BODY_FONT
                sr.font.size = Pt(REQUIRED_BODY_SIZE_PT)

            p = doc.add_paragraph(style=None)
            run = p.add_run(f"Q. {item.question}")
            run.font.name = REQUIRED_BODY_FONT
            run.font.size = Pt(REQUIRED_BODY_SIZE_PT)

            if item.options:
                for opt in item.options:
                    op = doc.add_paragraph()
                    orr = op.add_run(f"- {opt}")
                    orr.font.name = REQUIRED_BODY_FONT
                    orr.font.size = Pt(REQUIRED_BODY_SIZE_PT)


from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Pt


def insert_paragraph_after(paragraph, text=None, doc=None):
    """
    Safe helper to insert a paragraph after an existing paragraph.
    Uses the document's dominant body font if doc is provided,
    otherwise falls back to the configured body font.
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    if text:
        run = new_para.add_run(text)
        if doc is not None:
            dom_font, dom_size = _get_dominant_font(doc, is_heading=False)
            run.font.name = dom_font or REQUIRED_BODY_FONT
            run.font.size = Pt(dom_size or REQUIRED_BODY_SIZE_PT)
        else:
            run.font.name = REQUIRED_BODY_FONT
            run.font.size = Pt(REQUIRED_BODY_SIZE_PT)

    return new_para


def insert_generated_quizzes_after_units(doc, state):
    """
    Insert generated quiz only once per detected chapter/unit,
    strictly at the end of that unit.

    Important behavior:
    - quiz is inserted after unit.end_paragraph_index
    - never after subtopics/topics
    - skips units that already have a quiz
    - skips empty generated quiz sets
    """

    units = [
        u for u in getattr(state, "units", [])
        if (not u.quiz_present) and u.generated_quiz
    ]

    # process from bottom to top so paragraph shifts do not break positions
    units = sorted(units, key=lambda x: x.end_paragraph_index, reverse=True)

    for unit in units:
        insert_idx = unit.end_paragraph_index

        if insert_idx is None:
            continue

        if insert_idx < 0:
            continue

        if insert_idx >= len(doc.paragraphs):
            insert_idx = len(doc.paragraphs) - 1

        if insert_idx < 0 or len(doc.paragraphs) == 0:
            continue

        anchor_para = doc.paragraphs[insert_idx]

        # insert page break before quiz
        page_break_para = insert_paragraph_after(anchor_para)
        run = page_break_para.add_run()
        run.add_break(WD_BREAK.PAGE)

        # get dominant fonts for this document
        dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
        dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)
        h_font = dom_heading_font or REQUIRED_HEADING_FONT
        h_size = dom_heading_size or REQUIRED_HEADING_SIZE_PT
        b_font = dom_body_font or REQUIRED_BODY_FONT
        b_size = dom_body_size or REQUIRED_BODY_SIZE_PT

        # main quiz heading
        current_anchor = insert_paragraph_after(page_break_para, f"Quiz – {unit.title}")
        if current_anchor.runs:
            current_anchor.runs[0].bold = True
            current_anchor.runs[0].font.name = h_font
            current_anchor.runs[0].font.size = Pt(h_size)

        current_type = None

        for item in unit.generated_quiz:
            # section heading per question type
            if item.qtype != current_type:
                current_type = item.qtype
                current_anchor = insert_paragraph_after(current_anchor, current_type)
                if current_anchor.runs:
                    current_anchor.runs[0].bold = True
                    current_anchor.runs[0].font.name = b_font
                    current_anchor.runs[0].font.size = Pt(b_size)

            # question
            current_anchor = insert_paragraph_after(current_anchor, f"Q. {item.question}")
            if current_anchor.runs:
                current_anchor.runs[0].font.name = b_font
                current_anchor.runs[0].font.size = Pt(b_size)

            # options for MCQ
            if item.options:
                for opt in item.options:
                    current_anchor = insert_paragraph_after(current_anchor, f"- {opt}")
                    if current_anchor.runs:
                        current_anchor.runs[0].font.name = b_font
                        current_anchor.runs[0].font.size = Pt(b_size)

def insert_diagram_placeholders(doc: Document, state):
    """
    Adds small placeholder note near sections recommended for diagrams.
    Conservative: insert only one short note before the target paragraph if not already present.
    """
    recs = sorted(getattr(state, "diagram_recommendations", []), key=lambda x: x["paragraph_index"], reverse=True)

    for rec in recs:
        idx = rec.get("paragraph_index")
        if idx is None or idx < 0 or idx >= len(doc.paragraphs):
            continue

        para = doc.paragraphs[idx]

        # Avoid duplicates
        prev_text = ""
        if idx > 0:
            prev_text = doc.paragraphs[idx - 1].text.strip().lower()

        if "diagram recommended" in prev_text:
            continue

        note = insert_paragraph_after(
            para,
            "[Diagram Recommended: Add a flowchart/architecture/comparison visual here for better clarity.]"
        )
        if note.runs:
            note.runs[0].italic = True
            note.runs[0].font.name = REQUIRED_BODY_FONT
            note.runs[0].font.size = Pt(REQUIRED_BODY_SIZE_PT)

  
def insert_generated_diagrams(doc: Document, state):
    diagrams = sorted(getattr(state, "generated_diagrams", []), key=lambda x: x["paragraph_index"], reverse=True)

    for item in diagrams:
        img_path = item.get("image_path")
        para_idx = item.get("paragraph_index")

        if not img_path:
            continue
        if para_idx is None or para_idx < 0 or para_idx >= len(doc.paragraphs):
            continue

        anchor_para = doc.paragraphs[para_idx]

        caption_para = insert_paragraph_after(
            anchor_para,
            f"[Auto-generated {item.get('diagram_type', 'diagram')} diagram]"
        )
        if caption_para.runs:
            caption_para.runs[0].italic = True
            caption_para.runs[0].font.name = REQUIRED_BODY_FONT
            caption_para.runs[0].font.size = Pt(REQUIRED_BODY_SIZE_PT)

        image_para = insert_paragraph_after(caption_para)
        run = image_para.add_run()
        run.add_picture(img_path, width=Inches(DIAGRAM_IMAGE_WIDTH_INCHES))

def polish_doc_structure(doc: Document):
    # remove repeated blank paragraphs — but never remove paragraphs with images
    blank_paras = []
    prev_blank = False

    for para in doc.paragraphs:
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        is_blank = not (para.text and para.text.strip()) and not has_drawing
        if is_blank and prev_blank:
            blank_paras.append(para)
        prev_blank = is_blank

    for para in blank_paras:
        _remove_paragraph(para)

def insert_advanced_visuals(doc: Document, state):
    visuals = sorted(
        [v for v in getattr(state, "generated_visuals", []) if v.status == "success" and v.image_path],
        key=lambda x: x.paragraph_index,
        reverse=True,
    )

    for item in visuals:
        idx = item.paragraph_index
        if idx is None or idx < 0 or idx >= len(doc.paragraphs):
            continue

        anchor_para = doc.paragraphs[idx]

        label_para = insert_paragraph_after(
            anchor_para,
            f"[Auto-generated {item.visual_type.replace('_', ' ')} visual: {item.title}]"
        )
        if label_para.runs:
            label_para.runs[0].italic = True
            label_para.runs[0].font.name = REQUIRED_BODY_FONT
            label_para.runs[0].font.size = Pt(REQUIRED_BODY_SIZE_PT)

        image_para = insert_paragraph_after(label_para)
        run = image_para.add_run()
        run.add_picture(item.image_path, width=Inches(ADVANCED_VISUAL_WIDTH_INCHES))


def build_review_comments_doc(original_file: Path, state, output_path: Path):
    doc = copy_docx(original_file)
    add_inline_comments_to_doc(doc, state)
    add_ocr_summary_to_review_doc(doc, state)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def justify_body_paragraphs(doc: Document):
    """
    Apply justified alignment to all body paragraphs that don't already
    have an explicit alignment set. Skips:
    - Headings (they keep their own alignment)
    - Image paragraphs (centered images should stay centered)
    - Paragraphs explicitly set to CENTER, RIGHT, or other alignments
    """
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"

        # Skip headings
        if style_name.lower().startswith("heading"):
            continue

        # Skip image paragraphs
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue

        # Skip paragraphs with explicit non-justify alignment
        if para.alignment in (WD_ALIGN_PARAGRAPH.CENTER,
                               WD_ALIGN_PARAGRAPH.RIGHT,
                               WD_ALIGN_PARAGRAPH.DISTRIBUTE):
            continue

        # Apply justify to unset or left-aligned paragraphs
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def build_final_fixed_doc(original_file: Path, state, output_path: Path):
    doc = copy_docx(original_file)
    clean_doc_spacing(doc)
    remove_exact_duplicate_body_paragraphs(doc)
    apply_rewrites_to_doc(doc, state)
    add_missing_summary_headings(doc)

    # Find actual source URLs for images missing sources using Mistral vision
    from image_source_finder import find_sources_for_images
    image_sources = find_sources_for_images(state)

    # Start figure counter from highest existing figure number in the doc
    figure_counter = _get_next_figure_number(doc)
    caption_inserted_at = set()  # paragraphs that already got a caption this run
    source_inserted_at = set()   # paragraphs that already got a source this run

    for img in state.images:
        if img.paragraph_index is None:
            continue

        target_idx = img.nearest_caption_paragraph_index if img.nearest_caption_paragraph_index is not None else img.paragraph_index

        needs_caption = (not img.has_figure_number) and (target_idx not in caption_inserted_at)
        needs_source = (not img.has_source_link) and (target_idx not in source_inserted_at)

        if needs_caption or needs_source:
            figure_text = f"Figure {figure_counter}: Image description/caption to be finalized." if needs_caption else None

            # Use real source string if found (already formatted as "Source: ...")
            if needs_source:
                found_source = image_sources.get(img.filename)
                source_text = found_source if found_source else "Source: [Add source URL/reference here]"
            else:
                source_text = None

            _append_caption_then_source(doc, target_idx, figure_text, source_text)

            if needs_caption:
                caption_inserted_at.add(target_idx)
                figure_counter += 1
            if needs_source:
                source_inserted_at.add(target_idx)

    # quiz must be inserted once per chapter/unit only
    insert_generated_quizzes_after_units(doc, state)

    # keep your visuals/diagrams if present
    if hasattr(state, "generated_visuals"):
        insert_advanced_visuals(doc, state)
    elif hasattr(state, "generated_diagrams"):
        insert_generated_diagrams(doc, state)

    polish_doc_structure(doc)
    justify_body_paragraphs(doc)
    normalize_styles(doc)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))