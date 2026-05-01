from collections import Counter
from pathlib import Path
from typing import Optional
import re

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from docx.text.paragraph import Paragraph

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


# ─────────────────────────────────────────────────────────────────────────────
# Font helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_dominant_font(doc: Document, is_heading: bool):
    """
    Find the most commonly used font name and size for heading or body paragraphs.
    Skips TOC entries — they use a different style by design and should not
    influence the dominant font calculation.
    Falls back to Normal style definition if no run-level fonts are found.
    """
    font_names = []
    font_sizes = []

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        style_lower = style_name.lower()

        # Skip TOC entries — never include in dominant font calculation
        if style_lower.startswith("toc"):
            continue

        para_is_heading = style_lower.startswith("heading")
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

    # Fallback to style definition if no explicit run-level fonts found
    if not font_names and not is_heading:
        try:
            normal_style = doc.styles["Normal"]
            if normal_style.font and normal_style.font.name:
                font_names = [normal_style.font.name]
            if normal_style.font and normal_style.font.size:
                font_sizes = [normal_style.font.size.pt]
        except Exception:
            pass

    dominant_font = Counter(font_names).most_common(1)[0][0] if font_names else None
    dominant_size = Counter(font_sizes).most_common(1)[0][0] if font_sizes else None
    return dominant_font, dominant_size


# ─────────────────────────────────────────────────────────────────────────────
# Comment stripping
# ─────────────────────────────────────────────────────────────────────────────

def _strip_all_comments(doc: Document):
    """
    Remove ALL existing comments from a document.

    Final Fixed doc must be completely clean — zero comments.
    This strips:
    - Comments from previous manual reviewers (e.g. shobhit jaiswal)
    - Any comments carried over from the original document
    - Comment reference marks in paragraph runs

    Must be called immediately after copy_docx() in build_final_fixed_doc.
    """
    # Remove comment reference marks from all paragraphs
    for para in doc.paragraphs:
        for elem in para._element.xpath('.//*[local-name()="commentRangeStart"]'):
            elem.getparent().remove(elem)
        for elem in para._element.xpath('.//*[local-name()="commentRangeEnd"]'):
            elem.getparent().remove(elem)
        for elem in para._element.xpath('.//*[local-name()="commentReference"]'):
            elem.getparent().remove(elem)

    # Clear the comments XML part entirely
    try:
        part = doc.part
        if hasattr(part, '_comments_part') and part._comments_part is not None:
            comments_elem = part._comments_part._element
            if comments_elem is not None:
                # Remove all w:comment children
                for child in list(comments_elem):
                    comments_elem.remove(child)
    except Exception:
        pass

    # Also clear via relationship if accessible
    try:
        for rel in list(doc.part.rels.values()):
            if 'comments' in rel.reltype.lower():
                comments_part = rel.target_part
                root = comments_part._element
                for child in list(root):
                    root.remove(child)
                break
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# Style normalization
# ─────────────────────────────────────────────────────────────────────────────

def normalize_styles(doc: Document):
    """
    Fix only paragraphs whose font/size is inconsistent with the document's
    own dominant style. Does NOT force Arial — uses the document's own norm.
    Skips TOC entries entirely.
    """
    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)

    # Update TOC style definitions to use dominant body font
    toc_style_names = [
        "toc 1", "toc 2", "toc 3", "toc 4", "toc 5",
        "toc 6", "toc 7", "toc 8", "toc 9", "TOC Heading"
    ]
    target_toc_font = dom_body_font or REQUIRED_BODY_FONT
    for sn in toc_style_names:
        try:
            style = doc.styles[sn]
            if style.font:
                style.font.name = target_toc_font
        except KeyError:
            pass

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        style_lower = style_name.lower()

        # Never touch TOC paragraphs
        if style_lower.startswith("toc"):
            continue

        is_heading = style_lower.startswith("heading")
        target_font = dom_heading_font if is_heading else dom_body_font
        target_size = dom_heading_size if is_heading else dom_body_size

        if not target_font and not target_size:
            continue

        for run in para.runs:
            if not run.text.strip():
                continue

            run_font = run.font.name
            run_size = run.font.size.pt if run.font.size else None

            if run_font and target_font and run_font != target_font:
                run.font.name = target_font
            if run_size is not None and target_size is not None:
                if abs(run_size - target_size) > 0.5:
                    run.font.size = Pt(target_size)
            if is_heading:
                run.bold = True


# ─────────────────────────────────────────────────────────────────────────────
# Spacing cleanup
# ─────────────────────────────────────────────────────────────────────────────

def clean_doc_spacing(doc: Document):
    """
    Clean extra spacing in paragraphs without destroying inline formatting.
    Edits run text in-place — never uses para.clear() which destroys bold/italic/links.
    Uses document's own dominant font, not hardcoded Arial.
    Skips TOC entries and image paragraphs.
    """
    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        style_lower = style_name.lower()

        # Never touch TOC entries or image paragraphs
        if style_lower.startswith("toc"):
            continue
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue
        if not para.text or not para.text.strip():
            continue

        cleaned = clean_spacing(para.text)
        if cleaned == para.text.strip():
            continue

        is_heading = style_lower.startswith("heading")
        target_font = dom_heading_font if is_heading else dom_body_font
        target_size = dom_heading_size if is_heading else dom_body_size

        meaningful_runs = [r for r in para.runs if r.text.strip()]

        if len(meaningful_runs) == 1:
            run = meaningful_runs[0]
            run.text = cleaned
            if target_font and run.font.name and run.font.name != target_font:
                run.font.name = target_font
            if target_size and run.font.size:
                if abs(run.font.size.pt - target_size) > 0.5:
                    run.font.size = Pt(target_size)

        elif len(meaningful_runs) > 1:
            meaningful_runs[0].text = cleaned
            for run in meaningful_runs[1:]:
                run.text = ""
            run = meaningful_runs[0]
            if target_font and run.font.name and run.font.name != target_font:
                run.font.name = target_font
            if target_size and run.font.size:
                if abs(run.font.size.pt - target_size) > 0.5:
                    run.font.size = Pt(target_size)


# ─────────────────────────────────────────────────────────────────────────────
# Document helpers
# ─────────────────────────────────────────────────────────────────────────────

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
            for idx in [paragraph_index - distance, paragraph_index + distance]:
                if 0 <= idx < len(doc.paragraphs):
                    candidate = doc.paragraphs[idx]
                    if _paragraph_has_meaningful_text(candidate) or candidate.runs:
                        return candidate, idx

    for idx, p in enumerate(doc.paragraphs):
        if _paragraph_has_meaningful_text(p) or p.runs:
            return p, idx

    return doc.paragraphs[0], 0


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


def insert_paragraph_after(paragraph, text=None, doc=None):
    """Insert a new paragraph after the given paragraph."""
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


def _remove_paragraph(paragraph):
    p = paragraph._element
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


# ─────────────────────────────────────────────────────────────────────────────
# Issue tag mapping
# ─────────────────────────────────────────────────────────────────────────────

_ISSUE_TYPE_TAG = {
    "grammar":                          "[Grammar]",
    "spacing":                          "[Format]",
    "formatting_body":                  "[Format]",
    "formatting_heading":               "[Format]",
    "heading_hierarchy":                "[Structure]",
    "navigation_pane":                  "[Structure]",
    "repeated_heading":                 "[Structure]",
    "duplicate_paragraph":              "[Duplicate]",
    "near_duplicate_paragraph":         "[Duplicate]",
    "semantic_duplicate_paragraph":     "[Duplicate]",
    "possible_duplicate_topic":         "[Duplicate]",
    "image_caption_missing":            "[Image]",
    "image_source_missing":             "[Image]",
    "figure_number_missing":            "[Image]",
    "ocr_low_confidence":               "[OCR]",
    "content_accuracy":                 "[Accuracy]",
    "quiz_missing":                     "[Quiz]",
    "diagram_recommended":              "[Diagram]",
    "diagram_generated":                "[Diagram]",
    "diagram_generation_failed":        "[Diagram]",
    "advanced_visual_generated":        "[Visual]",
    "possible_ibm_similarity":          "[Plagiarism]",
    "possible_open_source_similarity":  "[Plagiarism]",
    "winston_scan_error":               "[Plagiarism]",
}


def _get_issue_tag(issue_type: str) -> str:
    return _ISSUE_TYPE_TAG.get(
        issue_type,
        f"[{issue_type.replace('_', ' ').title()}]"
    )


def _safe_comment_text(issue):
    tag = _get_issue_tag(issue.issue_type)
    return (
        f"{tag}\n"
        f"Severity: {issue.severity.upper()}\n"
        f"Location: {issue.location}\n"
        f"Comment: {issue.message}\n"
        f"Suggested Fix: {issue.suggested_fix}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Review Comments document builders
# ─────────────────────────────────────────────────────────────────────────────

def add_inline_comments_to_doc(doc: Document, state):
    """Add all pipeline review issues as Word comments in the Review doc."""
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
            try:
                doc.add_comment(
                    fallback_run,
                    text=_safe_comment_text(issue),
                    author=COMMENT_AUTHOR,
                    initials=COMMENT_INITIALS,
                )
            except Exception:
                pass


def add_ocr_summary_to_review_doc(doc: Document, state):
    # Internal pipeline info — not shown in Review Comments document
    pass


def _insert_review_summary(doc: Document, state) -> None:
    """
    Inserts a summary page at the very beginning of the Review Comments doc.
    Shows total issues, severity breakdown, top issue types, and units detected.
    """
    issues = getattr(state, "issues", [])
    units  = getattr(state, "units", [])

    severity_counts = Counter(i.severity for i in issues)
    type_counts     = Counter(i.issue_type for i in issues)

    body = doc.element.body

    def _prepend_para(text: str, bold: bool = False,
                      size_pt: int = None, space_after: int = 4) -> None:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        if size_pt:
            run.font.size = Pt(size_pt)
        pPr = p._p.get_or_add_pPr()
        spc = OxmlElement("w:spacing")
        spc.set(qn("w:after"), str(space_after * 20))
        pPr.append(spc)
        body.remove(p._p)
        body.insert(0, p._p)

    # Build reverse so final doc reads top-to-bottom

    # Page break after summary
    p_break = doc.add_paragraph()
    p_break.add_run().add_break(WD_BREAK.PAGE)
    body.remove(p_break._p)
    body.insert(0, p_break._p)

    # Units
    if units:
        for unit in reversed(units):
            quiz_note = ""
            if unit.quiz_present:
                quiz_note = "  ✓ quiz present"
            elif unit.generated_quiz:
                quiz_note = f"  ({len(unit.generated_quiz)} questions generated)"
            _prepend_para(f"  • {unit.title}{quiz_note}")
        _prepend_para("Units / Chapters Detected", bold=True)

    # Top issue types
    if type_counts:
        for issue_type, count in reversed(type_counts.most_common(8)):
            tag = _get_issue_tag(issue_type)
            _prepend_para(f"  {tag}  —  {count} issue(s)")
        _prepend_para("Top Issue Types", bold=True)

    # Severity breakdown
    for label in reversed(["info", "low", "medium", "high"]):
        count = severity_counts.get(label, 0)
        icon  = {"high": "🔴", "medium": "🟠", "low": "🟡", "info": "🔵"}.get(label, "⚪")
        _prepend_para(f"  {icon}  {label.upper()}: {count}")
    _prepend_para("Issues by Severity", bold=True)

    # Total
    _prepend_para(f"Total Issues Found: {len(issues)}", bold=True, size_pt=12)

    # Title
    _prepend_para("COURSEWARE REVIEW SUMMARY", bold=True, size_pt=16, space_after=8)


def build_review_comments_doc(original_file: Path, state, output_path: Path):
    """
    Build the Review Comments document.
    Copies the original, adds inline Word comments for every issue,
    and prepends a summary page.
    Does NOT strip existing comments — reviewer may want to see them.
    """
    doc = copy_docx(original_file)
    add_inline_comments_to_doc(doc, state)
    add_ocr_summary_to_review_doc(doc, state)
    _insert_review_summary(doc, state)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


# ─────────────────────────────────────────────────────────────────────────────
# Final Fixed document builders
# ─────────────────────────────────────────────────────────────────────────────

def _is_figure_caption(text: str) -> bool:
    if not text:
        return False
    return bool(re.search(r'^\s*(figure|fig\.?|fig)\s*\d+', text, flags=re.IGNORECASE))


def _get_next_figure_number(doc: Document) -> int:
    max_fig = 0
    for para in doc.paragraphs:
        m = re.search(r'figure\s*(\d+)', para.text, flags=re.IGNORECASE)
        if m:
            num = int(m.group(1))
            if num > max_fig:
                max_fig = num
    return max_fig + 1


def _append_caption_then_source(doc: Document, idx: int, figure_text: str, source_text: str):
    """Insert figure caption then source link after an image paragraph."""
    if idx is None or idx < 0 or idx >= len(doc.paragraphs):
        anchor = doc.paragraphs[-1] if doc.paragraphs else doc.add_paragraph()
    else:
        anchor = doc.paragraphs[idx]

    def _has_caption_nearby(para_idx):
        for check_idx in range(para_idx, min(para_idx + 4, len(doc.paragraphs))):
            if _is_figure_caption(doc.paragraphs[check_idx].text):
                return True
        return False

    caption_anchor = anchor

    if figure_text and not _has_caption_nearby(idx if idx is not None else 0):
        new_cap = insert_paragraph_after(caption_anchor)
        run = new_cap.add_run(figure_text)
        run.italic = True
        caption_anchor = new_cap

    if source_text:
        new_src = insert_paragraph_after(caption_anchor)
        run = new_src.add_run(source_text)
        run.italic = True


def add_missing_summary_headings(doc: Document):
    """
    Detect paragraphs that begin with summary/conclusion phrases and
    insert a Heading 3 before them if one is not already present.
    """
    SUMMARY_PATTERNS = [
        r'^in summary[,\s]', r'^to summarize[,\s]',
        r'^in conclusion[,\s]', r'^to conclude[,\s]',
        r'^summary[:\s]', r'^conclusion[:\s]',
        r'^overall[,\s]', r'^in brief[,\s]',
    ]

    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    h_font = dom_heading_font or REQUIRED_HEADING_FONT
    h_size = dom_heading_size or REQUIRED_HEADING_SIZE_PT

    to_insert = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip().lower()
        style = (para.style.name if para.style else "Normal").lower()
        if style.startswith("heading") or style.startswith("toc"):
            continue
        if any(re.match(p, text) for p in SUMMARY_PATTERNS):
            if i > 0:
                prev_style = (doc.paragraphs[i - 1].style.name if doc.paragraphs[i - 1].style else "").lower()
                if prev_style.startswith("heading"):
                    continue
            to_insert.append(i)

    for i in sorted(to_insert, reverse=True):
        anchor = doc.paragraphs[i]
        new_p = OxmlElement("w:p")
        anchor._p.addprevious(new_p)
        new_para = Paragraph(new_p, anchor._parent)
        try:
            new_para.style = doc.styles["Heading 3"]
        except KeyError:
            pass
        run = new_para.add_run("Summary")
        run.font.name = h_font
        run.font.size = Pt(h_size)
        run.bold = True


def remove_exact_duplicate_body_paragraphs(doc: Document):
    """
    Remove duplicate body paragraphs that appear within MAX_DUPLICATE_DISTANCE
    paragraphs of each other. Distant duplicates are likely intentional repetition.
    Never removes headings, TOC entries, or image paragraphs.
    """
    MAX_DUPLICATE_DISTANCE = 20
    seen = {}
    to_remove = []

    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else "Normal"
        style_lower = style_name.lower()
        if style_lower.startswith("heading") or style_lower.startswith("toc"):
            continue
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue
        norm = normalize_text(para.text).lower()
        if not norm or len(norm) < 80:
            continue
        if norm in seen:
            first_idx = seen[norm]
            if i - first_idx <= MAX_DUPLICATE_DISTANCE:
                to_remove.append(para)
            seen[norm] = i
        else:
            seen[norm] = i

    for para in to_remove:
        _remove_paragraph(para)


def apply_rewrites_to_doc(doc: Document, state):
    """
    Apply LLM rewrite suggestions in-place.
    Edits first run text only — preserves all inline formatting.
    Uses document's own dominant font, not hardcoded Arial.
    """
    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)

    for item in getattr(state, "rewrite_suggestions", []):
        if item.skipped or not item.applied:
            continue
        idx = item.paragraph_index
        if idx is None or idx < 0 or idx >= len(doc.paragraphs):
            continue

        para = doc.paragraphs[idx]
        has_image = para._element.xpath(
            './/*[local-name()="drawing"] | .//*[local-name()="pict"]'
        )
        if has_image:
            continue

        style_name = (para.style.name if para.style else "Normal").lower()
        is_heading = style_name.startswith("heading")
        target_font = dom_heading_font if is_heading else dom_body_font
        target_size = dom_heading_size if is_heading else dom_body_size

        meaningful_runs = [r for r in para.runs if r.text.strip()]

        if not meaningful_runs:
            run = para.add_run(item.rewritten_text)
            if target_font:
                run.font.name = target_font
            if target_size:
                run.font.size = Pt(target_size)
            if is_heading:
                run.bold = True
        else:
            meaningful_runs[0].text = item.rewritten_text
            for run in meaningful_runs[1:]:
                run.text = ""
            first_run = meaningful_runs[0]
            if target_font and first_run.font.name and first_run.font.name != target_font:
                first_run.font.name = target_font
            if target_size and first_run.font.size:
                if abs(first_run.font.size.pt - target_size) > 0.5:
                    first_run.font.size = Pt(target_size)
            if is_heading:
                first_run.bold = True


def _find_unit_end_in_doc(doc: Document, unit_title: str) -> Optional[int]:
    unit_title_norm = unit_title.strip().lower()
    unit_heading_idx = None
    unit_level = None

    for i, para in enumerate(doc.paragraphs):
        style_name = (para.style.name if para.style else "").lower()
        is_word_heading = style_name.startswith("heading")
        is_bold_heading = (
            bool(para.text.strip())
            and len(para.text.strip().split()) <= 12
            and bool(para.runs)
            and all(r.bold for r in para.runs if r.text.strip())
        )
        if not (is_word_heading or is_bold_heading):
            continue
        if para.text.strip().lower() == unit_title_norm:
            unit_heading_idx = i
            parts = style_name.split()
            if len(parts) >= 2 and parts[1].isdigit():
                unit_level = int(parts[1])
            break

    if unit_heading_idx is None:
        print(f"[QUIZ INSERT] Could not find heading '{unit_title}' in live doc — using end_paragraph_index fallback")
        return None

    for i in range(unit_heading_idx + 1, len(doc.paragraphs)):
        para_i = doc.paragraphs[i]
        style_name = (para_i.style.name if para_i.style else "").lower()
        is_word_heading = style_name.startswith("heading")
        is_bold_heading = (
            bool(para_i.text.strip())
            and len(para_i.text.strip().split()) <= 12
            and bool(para_i.runs)
            and all(r.bold for r in para_i.runs if r.text.strip())
        )
        if is_word_heading:
            parts = style_name.split()
            if len(parts) >= 2 and parts[1].isdigit():
                level = int(parts[1])
                if unit_level is None or level <= unit_level:
                    return i - 1
        elif is_bold_heading:
            return i - 1

    return len(doc.paragraphs) - 1

def insert_generated_quizzes_after_units(doc: Document, state):
    """
    Insert generated quiz at the end of each unit that is missing one.

    DRIFT-SAFE: Uses heading text match to find position in the live document
    instead of the pre-recorded end_paragraph_index which drifts after rewrites
    and caption insertions.

    Processing order: bottom-to-top so earlier insertions don't shift later positions.
    Font: uses document's own dominant font, not hardcoded Arial.
    """
    units = [
        u for u in getattr(state, "units", [])
        if (not u.quiz_present) and u.generated_quiz
    ]
    if not units:
        return

    # Resolve all positions before any insertion
    resolved = []
    for unit in units:
        insert_idx = _find_unit_end_in_doc(doc, unit.title)
        if insert_idx is None:
            insert_idx = unit.end_paragraph_index  # fallback
        if insert_idx is None:
            continue
        insert_idx = max(0, min(insert_idx, len(doc.paragraphs) - 1))
        resolved.append((unit, insert_idx))

    # Sort bottom-to-top
    resolved.sort(key=lambda x: x[1], reverse=True)

    # Cache fonts once — expensive to compute per unit on large docs
    dom_heading_font, dom_heading_size = _get_dominant_font(doc, is_heading=True)
    dom_body_font, dom_body_size = _get_dominant_font(doc, is_heading=False)
    h_font = dom_heading_font or REQUIRED_HEADING_FONT
    h_size = dom_heading_size or REQUIRED_HEADING_SIZE_PT
    b_font = dom_body_font or REQUIRED_BODY_FONT
    b_size = dom_body_size or REQUIRED_BODY_SIZE_PT

    for unit, insert_idx in resolved:
        anchor_para = doc.paragraphs[insert_idx]

        # Page break before quiz
        page_break_para = insert_paragraph_after(anchor_para)
        page_break_para.add_run().add_break(WD_BREAK.PAGE)

        # Quiz heading
        current_anchor = insert_paragraph_after(page_break_para, f"Quiz – {unit.title}")
        if current_anchor.runs:
            current_anchor.runs[0].bold = True
            current_anchor.runs[0].font.name = h_font
            current_anchor.runs[0].font.size = Pt(h_size)

        current_type = None

        for item in unit.generated_quiz:
            # Section subheading per question type
            if item.qtype != current_type:
                current_type = item.qtype
                current_anchor = insert_paragraph_after(current_anchor, current_type)
                if current_anchor.runs:
                    current_anchor.runs[0].bold = True
                    current_anchor.runs[0].font.name = b_font
                    current_anchor.runs[0].font.size = Pt(b_size)

            # Question
            current_anchor = insert_paragraph_after(current_anchor, f"Q. {item.question}")
            if current_anchor.runs:
                current_anchor.runs[0].font.name = b_font
                current_anchor.runs[0].font.size = Pt(b_size)

            # MCQ options
            if item.options:
                for opt in item.options:
                    current_anchor = insert_paragraph_after(current_anchor, f"- {opt}")
                    if current_anchor.runs:
                        current_anchor.runs[0].font.name = b_font
                        current_anchor.runs[0].font.size = Pt(b_size)


def insert_generated_diagrams(doc: Document, state):
    diagrams = sorted(
        getattr(state, "generated_diagrams", []),
        key=lambda x: x["paragraph_index"],
        reverse=True
    )
    for item in diagrams:
        img_path = item.get("image_path")
        para_idx = item.get("paragraph_index")
        if not img_path or para_idx is None:
            continue
        if para_idx < 0 or para_idx >= len(doc.paragraphs):
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
        image_para.add_run().add_picture(img_path, width=Inches(DIAGRAM_IMAGE_WIDTH_INCHES))


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
        image_para.add_run().add_picture(item.image_path, width=Inches(ADVANCED_VISUAL_WIDTH_INCHES))


def polish_doc_structure(doc: Document):
    """Remove consecutive blank paragraphs. Never removes image paragraphs."""
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


def justify_body_paragraphs(doc: Document):
    """Apply justified alignment to body paragraphs. Skips headings, images, TOC, centred text."""
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        style_lower = style_name.lower()
        if style_lower.startswith("heading") or style_lower.startswith("toc"):
            continue
        has_drawing = bool(para._element.xpath('.//*[local-name()="drawing"]'))
        if has_drawing:
            continue
        if para.alignment in (
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.DISTRIBUTE,
        ):
            continue
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def build_final_fixed_doc(original_file: Path, state, output_path: Path):
    """
    Build the Final Fixed document.

    Order of operations:
    1. Copy original
    2. Strip ALL existing comments (pipeline must not add any — Final Fixed is clean)
    3. Clean spacing (in-place, preserving formatting)
    4. Remove close-range duplicate paragraphs
    5. Apply LLM rewrites (in-place)
    6. Add missing summary headings
    7. Add missing image captions and source links
    8. Insert generated quizzes at unit ends (drift-safe)
    9. Insert advanced visuals / diagrams
    10. Polish structure (remove excess blanks)
    11. Justify body paragraphs
    12. Normalize styles using document's own dominant font
    """
    doc = copy_docx(original_file)

    # ── Step 2: Strip ALL comments — Final Fixed must be completely clean ──
    _strip_all_comments(doc)

    # ── Step 3: Spacing ───────────────────────────────────────────────────
    clean_doc_spacing(doc)

    # ── Step 4: Remove close-range duplicate paragraphs ───────────────────
    remove_exact_duplicate_body_paragraphs(doc)

    # ── Step 5: LLM rewrites ──────────────────────────────────────────────
    apply_rewrites_to_doc(doc, state)

    # ── Step 6: Summary headings ──────────────────────────────────────────
    add_missing_summary_headings(doc)

    # ── Step 7: Image captions and source links ───────────────────────────
    from image_source_finder import find_sources_for_images
    image_sources = find_sources_for_images(state)

    figure_counter = _get_next_figure_number(doc)
    caption_inserted_at = set()
    source_inserted_at = set()

    for img in state.images:
        if img.paragraph_index is None:
            continue
        target_idx = (
            img.nearest_caption_paragraph_index
            if img.nearest_caption_paragraph_index is not None
            else img.paragraph_index
        )
        needs_caption = (not img.has_figure_number) and (target_idx not in caption_inserted_at)
        needs_source  = (not img.has_source_link)    and (target_idx not in source_inserted_at)

        if needs_caption or needs_source:
            figure_text = (
                f"Figure {figure_counter}: Image description/caption to be finalized."
                if needs_caption else None
            )
            if needs_source:
                found_source = image_sources.get(img.filename)
                source_text = found_source or "Source: [Add source URL/reference here]"
            else:
                source_text = None

            _append_caption_then_source(doc, target_idx, figure_text, source_text)

            if needs_caption:
                caption_inserted_at.add(target_idx)
                figure_counter += 1
            if needs_source:
                source_inserted_at.add(target_idx)

    # ── Step 8: Quizzes (drift-safe) ──────────────────────────────────────
    insert_generated_quizzes_after_units(doc, state)

    # ── Step 9: Visuals / diagrams ────────────────────────────────────────
    if getattr(state, "generated_visuals", None):
        insert_advanced_visuals(doc, state)
    elif getattr(state, "generated_diagrams", None):
        insert_generated_diagrams(doc, state)

    # ── Step 10: Polish ───────────────────────────────────────────────────
    polish_doc_structure(doc)

    # ── Step 11: Justify ──────────────────────────────────────────────────
    justify_body_paragraphs(doc)

    # ── Step 12: Normalize fonts using document's own dominant style ──────
    normalize_styles(doc)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
