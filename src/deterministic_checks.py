from collections import Counter
from typing import List, Optional

from config import (
    REQUIRED_HEADING_FONT,
    REQUIRED_HEADING_SIZE_PT,
    REQUIRED_HEADING_BOLD,
    REQUIRED_BODY_FONT,
    REQUIRED_BODY_SIZE_PT,
    MIN_TEXT_CHARS_FOR_USEFUL_OCR,
)
from models import DocumentState, ReviewIssue
from utils import clean_spacing


def most_common_or_none(values: List):
    if not values:
        return None
    counter = Counter(v for v in values if v is not None)
    if not counter:
        return None
    return counter.most_common(1)[0][0]


def add_issue(state: DocumentState, issue_type: str, severity: str,
              paragraph_index: Optional[int], location: str,
              message: str, suggested_fix: str):
    state.issues.append(
        ReviewIssue(
            issue_type=issue_type,
            severity=severity,
            paragraph_index=paragraph_index,
            location=location,
            message=message,
            suggested_fix=suggested_fix,
        )
    )


def _image_para_index(state: DocumentState, image_filename: str):
    for img in state.images:
        if img.filename == image_filename:
            return img.paragraph_index
    return None


def _dominant_font(paragraphs, is_heading_filter: bool):
    """
    Return the most commonly used font name and size across the
    heading (or body) paragraphs in the document.
    Used to determine what the document's own standard is before
    flagging inconsistencies.
    """
    font_names = []
    font_sizes = []
    for para in paragraphs:
        if para.is_heading != is_heading_filter:
            continue
        if not is_heading_filter and not para.text.strip():
            continue
        fn = most_common_or_none(para.font_names)
        fs = most_common_or_none(para.font_sizes)
        if fn:
            font_names.append(fn)
        if fs:
            font_sizes.append(fs)

    dominant_font = most_common_or_none(font_names)
    dominant_size = most_common_or_none(font_sizes)
    return dominant_font, dominant_size


def check_heading_styles(state: DocumentState):
    """
    Flag headings that are inconsistent with the dominant heading style
    used across the document. Also flag bold inconsistency.
    Only paragraphs that deviate from the document's own norm are flagged.
    """
    dominant_font, dominant_size = _dominant_font(state.paragraphs, is_heading_filter=True)

    bold_values = []
    for para in state.paragraphs:
        if not para.is_heading:
            continue
        b = most_common_or_none(para.bold_flags)
        if b is not None:
            bold_values.append(b)
    dominant_bold = most_common_or_none(bold_values)

    for para in state.paragraphs:
        if not para.is_heading:
            continue

        font_name = most_common_or_none(para.font_names)
        font_size = most_common_or_none(para.font_sizes)
        bold_flag = most_common_or_none(para.bold_flags)

        problems = []

        if font_name and dominant_font and font_name != dominant_font:
            problems.append(f"inconsistent font: found '{font_name}', document norm is '{dominant_font}'")

        if font_size and dominant_size and abs(font_size - dominant_size) > 0.5:
            problems.append(f"inconsistent size: found {font_size} pt, document norm is {dominant_size} pt")

        if bold_flag is not None and dominant_bold is not None and bold_flag != dominant_bold:
            problems.append(f"inconsistent bold: found {bold_flag}, document norm is {dominant_bold}")

        if problems:
            add_issue(
                state=state,
                issue_type="formatting_heading",
                severity="medium",
                paragraph_index=para.index,
                location=f"Paragraph {para.index} | {para.style_name}",
                message="Heading formatting inconsistency: " + "; ".join(problems),
                suggested_fix=f"Apply the document's dominant heading style: '{dominant_font}', {dominant_size} pt, bold={dominant_bold}.",
            )


def check_body_styles(state: DocumentState):
    """
    Flag body paragraphs that are inconsistent with the dominant body style
    used across the document.
    Only paragraphs that deviate from the document's own norm are flagged.
    """
    dominant_font, dominant_size = _dominant_font(state.paragraphs, is_heading_filter=False)

    for para in state.paragraphs:
        if para.is_heading:
            continue
        if not para.text.strip():
            continue

        font_name = most_common_or_none(para.font_names)
        font_size = most_common_or_none(para.font_sizes)

        problems = []

        if font_name and dominant_font and font_name != dominant_font:
            problems.append(f"inconsistent font: found '{font_name}', document norm is '{dominant_font}'")

        if font_size and dominant_size and abs(font_size - dominant_size) > 0.5:
            problems.append(f"inconsistent size: found {font_size} pt, document norm is {dominant_size} pt")

        if problems:
            add_issue(
                state=state,
                issue_type="formatting_body",
                severity="low",
                paragraph_index=para.index,
                location=f"Paragraph {para.index} | {para.style_name}",
                message="Body formatting inconsistency: " + "; ".join(problems),
                suggested_fix=f"Apply the document's dominant body style: '{dominant_font}', {dominant_size} pt.",
            )


def check_heading_hierarchy(state: DocumentState):
    """
    Checks for structural heading issues that would break the Navigation Pane:
    1. Document has no H1 at all (no top-level anchor)
    2. Level skips (H1 -> H3, skipping H2)
    3. Duplicate consecutive headings at the same level with no content between them
    """
    headings = [p for p in state.paragraphs if p.is_heading and p.heading_level is not None]

    if not headings:
        return

    # Check 1: No H1 present
    has_h1 = any(h.heading_level == 1 for h in headings)
    if not has_h1:
        add_issue(
            state=state,
            issue_type="heading_hierarchy",
            severity="high",
            paragraph_index=headings[0].index,
            location="Document level",
            message="Document has no Heading 1. Navigation Pane requires at least one top-level (H1) heading.",
            suggested_fix="Promote the highest-level heading to Heading 1.",
        )

    # Check 2: Level skips
    prev_level = None
    for para in headings:
        curr = para.heading_level
        if prev_level is not None and curr > prev_level + 1:
            add_issue(
                state=state,
                issue_type="heading_hierarchy",
                severity="high",
                paragraph_index=para.index,
                location=f"Paragraph {para.index} | {para.style_name}",
                message=(
                    f"Heading level skips from H{prev_level} to H{curr}. "
                    f"Navigation Pane will show a broken hierarchy."
                ),
                suggested_fix=(
                    f"Change this heading to H{prev_level + 1} or insert a missing "
                    f"H{prev_level + 1} section before it."
                ),
            )
        prev_level = curr

    # Check 3: Consecutive same-level headings with no body content between them
    prev_heading = None
    for para in state.paragraphs:
        if para.is_heading:
            if (
                prev_heading is not None
                and para.heading_level == prev_heading.heading_level
            ):
                # Check if there is any non-empty body paragraph between them
                between = state.paragraphs[prev_heading.index + 1: para.index]
                has_content = any(
                    not p.is_heading and p.text.strip()
                    for p in between
                )
                if not has_content:
                    add_issue(
                        state=state,
                        issue_type="heading_hierarchy",
                        severity="medium",
                        paragraph_index=para.index,
                        location=f"Paragraph {para.index} | {para.style_name}",
                        message=(
                            f"Two consecutive H{para.heading_level} headings with no content "
                            f"between them (Paragraph {prev_heading.index} and {para.index}). "
                            f"Navigation Pane will show empty sections."
                        ),
                        suggested_fix="Add content under the first heading or merge the two headings.",
                    )
            prev_heading = para


def check_navigation_compatibility(state: DocumentState):
    """
    Checks overall Navigation Pane health:
    - No headings at all
    - Only one heading level used in a long document (flat structure)
    """
    headings = [p for p in state.paragraphs if p.is_heading]
    heading_count = len(headings)

    if heading_count == 0:
        add_issue(
            state=state,
            issue_type="navigation_pane",
            severity="high",
            paragraph_index=0 if state.paragraphs else None,
            location="Document level",
            message="No heading styles detected. Navigation Pane will be empty and unusable.",
            suggested_fix="Apply Word heading styles (Heading 1, Heading 2, etc.) to all section titles.",
        )
        return

    # Warn if a long document uses only one heading level — flat structure
    levels_used = set(p.heading_level for p in headings if p.heading_level is not None)
    total_paras = len(state.paragraphs)
    if len(levels_used) == 1 and total_paras > 30:
        add_issue(
            state=state,
            issue_type="navigation_pane",
            severity="medium",
            paragraph_index=headings[0].index,
            location="Document level",
            message=(
                f"Document uses only one heading level (H{list(levels_used)[0]}) across "
                f"{total_paras} paragraphs. Navigation Pane will not allow users to navigate "
                f"sub-topics."
            ),
            suggested_fix="Add sub-headings (Heading 2, Heading 3) to break the document into navigable sections.",
        )


def check_spacing_issues(state: DocumentState):
    for para in state.paragraphs:
        if para.has_extra_spacing_issue:
            cleaned = clean_spacing(para.text)
            add_issue(
                state=state,
                issue_type="spacing",
                severity="low",
                paragraph_index=para.index,
                location=f"Paragraph {para.index}",
                message="Extra spacing or spacing before punctuation detected.",
                suggested_fix=f"Suggested cleaned text: {cleaned}",
            )


def check_grammar_spelling(state: DocumentState):
    """
    Flag paragraphs that are likely to contain grammar or spelling issues
    so the LLM rewrite pipeline can correct them.
    Uses simple heuristics — the actual correction is done by Mistral.
    """
    # Words that start with a vowel letter but have a consonant sound.
    # These are correctly preceded by 'a' not 'an', so must be excluded.
    _consonant_sound_vowel_words = (
        "user|users|use|used|using|usable|usual|usually|utility|utilities|"
        "unique|unit|units|uniform|uniforms|union|unions|universal|universe|"
        "university|url|uri|uri|usher|ushers|"
        "european|europe|"
        "one|once|"
    )

    grammar_patterns = [
        # double spaces mid-sentence
        (r"  +", "Double spaces detected."),
        # 'a' before vowel sound — excludes words like 'user', 'unique', 'uniform'
        (
            r"\ba\s+(?!(?:" + _consonant_sound_vowel_words + r")\b)[aeiouAEIOU]\w{2,}",
            "Possible grammar issue: 'a' used before a vowel sound (should be 'an')."
        ),
        # repeated words (the the, is is)
        (r"\b(\w+)\s+\1\b", "Repeated word detected."),
        # sentence not starting with capital
        (r"(?<=[.!?]\s)[a-z]", "Sentence may not start with a capital letter."),
    ]

    import re
    for para in state.paragraphs:
        if para.is_heading:
            continue
        if getattr(para, "is_code_like", False):
            continue
        text = para.text or ""
        if len(text.strip()) < 20:
            continue

        for pattern, message in grammar_patterns:
            if re.search(pattern, text):
                add_issue(
                    state=state,
                    issue_type="grammar",
                    severity="low",
                    paragraph_index=para.index,
                    location=f"Paragraph {para.index}",
                    message=message,
                    suggested_fix="Review and correct grammar, spelling, and sentence structure.",
                )
                break  # one issue per paragraph is enough to trigger rewrite


def check_image_presence(state: DocumentState):
    # Intentionally empty — image presence is validated through OCR and
    # caption/source checks. The image_review_pending comment is noise.
    pass


def check_ocr_results(state: DocumentState):
    if not state.ocr_results:
        return

    for result in state.ocr_results:
        target_para_idx = _image_para_index(state, result.image_filename)

        # Only flag actionable OCR issues — low confidence means the text
        # extraction may be unreliable and needs manual review
        if result.status == "low_confidence":
            add_issue(
                state=state,
                issue_type="ocr_low_confidence",
                severity="medium",
                paragraph_index=target_para_idx,
                location=f"Image file: {result.image_filename}",
                message=f"Image text extraction confidence is low ({result.avg_confidence}). Content may be unreliable.",
                suggested_fix="Manually verify the image content and ensure it is clearly readable.",
            )


def check_image_caption_and_source(state: DocumentState):
    for img in state.images:
        # Skip images we couldn't locate in the document
        if img.paragraph_index is None:
            continue

        # Determine anchor paragraph - prefer caption location,
        # but never anchor on a heading
        raw_idx = (
            img.nearest_caption_paragraph_index
            if img.nearest_caption_paragraph_index is not None
            else img.paragraph_index
        )

        target_idx = raw_idx
        if raw_idx is not None and raw_idx < len(state.paragraphs):
            if state.paragraphs[raw_idx].is_heading:
                for search_idx in range(raw_idx + 1, min(raw_idx + 5, len(state.paragraphs))):
                    if not state.paragraphs[search_idx].is_heading:
                        target_idx = search_idx
                        break

        if not img.nearest_caption_text:
            add_issue(
                state=state,
                issue_type="image_caption_missing",
                severity="medium",
                paragraph_index=target_idx,
                location=f"Image file: {img.filename}",
                message="No nearby caption-like text found for this image.",
                suggested_fix="Add a caption near the image and include a figure number where applicable.",
            )

        if img.has_figure_number is False:
            add_issue(
                state=state,
                issue_type="figure_number_missing",
                severity="medium",
                paragraph_index=target_idx,
                location=f"Image file: {img.filename}",
                message="Image caption does not contain a figure number.",
                suggested_fix="Add a figure number such as 'Figure 1:' before the caption text.",
            )

        if img.has_source_link is False:
            add_issue(
                state=state,
                issue_type="image_source_missing",
                severity="high",
                paragraph_index=target_idx,
                location=f"Image file: {img.filename}",
                message="No nearby image source link/reference found.",
                suggested_fix="Add an image source URL or a clearly labeled source/reference line near the figure.",
            )


def check_duplicates(state: DocumentState):
    duplicate_findings = getattr(state, "duplicate_findings", [])
    for finding in duplicate_findings:
        ftype = finding.get("type")
        dup_idx = finding.get("duplicate_index")
        keep_idx = finding.get("keep_index")

        if ftype == "exact_duplicate_paragraph":
            add_issue(
                state=state,
                issue_type="duplicate_paragraph",
                severity="medium",
                paragraph_index=dup_idx,
                location=f"Paragraph {dup_idx}",
                message=f"This paragraph appears to exactly duplicate Paragraph {keep_idx}.",
                suggested_fix="Remove the repeated paragraph or merge the repeated content.",
            )
        elif ftype == "near_duplicate_paragraph":
            sim = finding.get("similarity")
            add_issue(
                state=state,
                issue_type="near_duplicate_paragraph",
                severity="medium",
                paragraph_index=dup_idx,
                location=f"Paragraph {dup_idx}",
                message=f"This paragraph is highly similar to Paragraph {keep_idx} (similarity: {sim}).",
                suggested_fix="Review and remove redundant wording or merge both paragraphs.",
            )
        elif ftype == "repeated_heading":
            add_issue(
                state=state,
                issue_type="repeated_heading",
                severity="medium",
                paragraph_index=dup_idx,
                location=f"Paragraph {dup_idx}",
                message=f"This heading appears to repeat the heading at Paragraph {keep_idx}.",
                suggested_fix="Check whether this is a duplicated topic and consolidate if needed.",
            )
        elif ftype == "possible_duplicate_topic":
            dups = finding.get("duplicate_indexes", [])
            add_issue(
                state=state,
                issue_type="possible_duplicate_topic",
                severity="high",
                paragraph_index=keep_idx,
                location=f"Paragraph {keep_idx}",
                message=f"This topic may be duplicated in paragraphs: {dups}.",
                suggested_fix="Review the duplicated topic sections and keep only the most complete version.",
            )


def check_retrieval_similarity(state: DocumentState):
    retrieval_findings = getattr(state, "retrieval_findings", [])
    for finding in retrieval_findings:
        para_idx = finding.get("paragraph_index")
        ftype = finding.get("type")

        if ftype == "winston_scan_error":
            add_issue(
                state=state,
                issue_type="winston_scan_error",
                severity="low",
                paragraph_index=para_idx,
                location=f"Paragraph {para_idx}",
                message=f"Plagiarism scan could not be completed for this paragraph. Error: {finding.get('error')}",
                suggested_fix="Retry the scan later or review this paragraph manually.",
            )
            continue

        if ftype == "winston_plagiarism_flag":
            score = finding.get("score", 0)
            source_title = finding.get("best_source_title", "")
            source_url = finding.get("best_source_url", "")
            total_plagiarism_words = finding.get("total_plagiarism_words", 0)
            source_count = finding.get("source_count", 0)
            best_source_is_excluded = finding.get("best_source_is_excluded", False)

            if best_source_is_excluded:
                add_issue(
                    state=state,
                    issue_type="possible_ibm_similarity",
                    severity="info",
                    paragraph_index=para_idx,
                    location=f"Paragraph {para_idx}",
                    message=f"Possible overlap with an excluded source/domain detected. Score: {score}. Sources: {source_count}. Matched words: {total_plagiarism_words}. Best source: {source_title} | {source_url}",
                    suggested_fix="Review this paragraph for originality. If the source is IBM-related, handle it according to your IBM exception policy.",
                )
            else:
                add_issue(
                    state=state,
                    issue_type="possible_open_source_similarity",
                    severity="high" if score >= 35 else "medium",
                    paragraph_index=para_idx,
                    location=f"Paragraph {para_idx}",
                    message=f"Possible overlap with public/open-source content detected. Score: {score}. Sources: {source_count}. Matched words: {total_plagiarism_words}. Best source: {source_title} | {source_url}",
                    suggested_fix="Rephrase this paragraph to ensure originality before finalizing the document.",
                )

def check_missing_quizzes(state: DocumentState):
    for unit in getattr(state, "units", []):
        if not unit.quiz_present:
            add_issue(
                state=state,
                issue_type="quiz_missing",
                severity="medium",
                paragraph_index=unit.heading_paragraph_index,
                location=f"Unit: {unit.title}",
                message="Quiz section appears to be missing for this unit.",
                suggested_fix="Add 15 quiz questions: 3 MCQ, 3 Fill in the Blanks, 3 True/False, 3 Two-mark, 3 Four-mark.",
            )

def check_diagram_recommendations(state: DocumentState):
    for item in getattr(state, "diagram_recommendations", []):
        para_idx = item.get("paragraph_index")
        score = item.get("score", 0)

        add_issue(
            state=state,
            issue_type="diagram_recommended",
            severity="medium",
            paragraph_index=para_idx,
            location=f"Paragraph {para_idx}",
            message=f"This section may be clearer if supplemented with a diagram or visual flow representation (score: {score}).",
            suggested_fix="Add a process flow, architecture diagram, lifecycle diagram, or comparison visual as appropriate.",
        )

def check_generated_diagrams(state: DocumentState):
    for item in getattr(state, "generated_diagrams", []):
        para_idx = item.get("paragraph_index")
        path = item.get("image_path")
        diag_type = item.get("diagram_type", "diagram")

        if path:
            add_issue(
                state=state,
                issue_type="diagram_generated",
                severity="info",
                paragraph_index=para_idx,
                location=f"Paragraph {para_idx}",
                message=f"A deterministic {diag_type} diagram was generated for this section.",
                suggested_fix="Verify that the inserted diagram accurately represents the section content.",
            )
        else:
            add_issue(
                state=state,
                issue_type="diagram_generation_failed",
                severity="low",
                paragraph_index=para_idx,
                location=f"Paragraph {para_idx}",
                message=f"Diagram generation failed for this section. Error: {item.get('error', 'Unknown error')}",
                suggested_fix="Manually create a diagram or retry generation later.",
            )

def check_table_findings(state: DocumentState):
    # Build a lookup: table_index -> paragraph_index
    table_para_map = {
        t.table_index: t.paragraph_index
        for t in getattr(state, "tables", [])
        if t.paragraph_index is not None
    }

    for item in getattr(state, "table_findings", []):
        t_idx = item.get("table_index")
        para_idx = table_para_map.get(t_idx, 0)

        add_issue(
            state=state,
            issue_type=item["type"],
            severity=item.get("severity", "low"),
            paragraph_index=para_idx,
            location=f"Table {t_idx} (near paragraph {para_idx})",
            message=item.get("message", "Table issue detected."),
            suggested_fix="Review the table structure, repeated rows, sparse cells, and readability.",
        )


def check_code_findings(state: DocumentState):
    for item in getattr(state, "code_findings", []):
        add_issue(
            state=state,
            issue_type=item["type"],
            severity=item.get("severity", "low"),
            paragraph_index=item.get("paragraph_index"),
            location=f"Paragraph {item.get('paragraph_index')}",
            message=item.get("message", "Code formatting issue detected."),
            suggested_fix="Preserve code content, but apply better code-block formatting and indentation consistency.",
        )

def check_visual_specs(state: DocumentState):
    # Visual spec creation is internal pipeline info — not shown to reviewer
    pass


def check_generated_visuals(state: DocumentState):
    # Only flag successfully generated visuals — failures are system issues
    # not content issues and should not appear in the review document
    for item in getattr(state, "generated_visuals", []):
        if item.status == "success":
            add_issue(
                state=state,
                issue_type="advanced_visual_generated",
                severity="info",
                paragraph_index=item.paragraph_index,
                location=f"Paragraph {item.paragraph_index}",
                message=f"An advanced {item.visual_type} visual was generated for this section.",
                suggested_fix="Review whether the generated visual explains the concept accurately and clearly.",
            )


def check_accuracy_findings(state: DocumentState):
    """
    Converts accuracy findings from run_accuracy_check() into
    ReviewIssue comments anchored to the relevant paragraph.
    Multiple findings on the same paragraph are merged into one comment.
    """
    from collections import defaultdict
    grouped = defaultdict(list)
    for finding in getattr(state, "accuracy_findings", []):
        grouped[finding.paragraph_index].append(finding)

    for para_idx, findings in grouped.items():
        if len(findings) == 1:
            f = findings[0]
            message = (
                f"Possible inaccuracy detected: {f.issue_description} "
                f"Flagged text: \"{f.flagged_text}\""
            )
            suggested_fix = f.suggestion
        else:
            # Merge all findings into one comment
            lines = []
            fixes = []
            for i, f in enumerate(findings, 1):
                lines.append(f"{i}. {f.issue_description} Flagged: \"{f.flagged_text}\"")
                fixes.append(f"{i}. {f.suggestion}")
            message = "Multiple possible inaccuracies detected:\n" + "\n".join(lines)
            suggested_fix = "\n".join(fixes)

        add_issue(
            state=state,
            issue_type="content_accuracy",
            severity="high",
            paragraph_index=para_idx,
            location=f"Unit: {findings[0].unit_title} | Paragraph {para_idx}",
            message=message,
            suggested_fix=suggested_fix,
        )


def run_all_step1_checks(state: DocumentState) -> DocumentState:
    """Step 1: formatting, hierarchy, spacing. No OCR or image metadata yet."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    return state


def run_all_checks_with_ocr(state: DocumentState) -> DocumentState:
    """Step 2: adds OCR result checks on top of step 1."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    return state


def run_all_checks_step3(state: DocumentState) -> DocumentState:
    """Step 3: adds image caption and source checks."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    return state


def run_all_checks_step4(state: DocumentState) -> DocumentState:
    """Step 4: adds duplicate detection."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    return state


def run_all_checks_step5(state: DocumentState) -> DocumentState:
    """Step 5 & 6: adds retrieval/Winston similarity."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    return state


def run_all_checks_step7(state: DocumentState) -> DocumentState:
    """Step 7: adds missing quiz checks and accuracy findings."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    check_accuracy_findings(state)
    check_missing_quizzes(state)
    return state


def run_all_checks_step8(state: DocumentState) -> DocumentState:
    """Step 8: adds diagram recommendation checks."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    check_accuracy_findings(state)
    check_missing_quizzes(state)
    check_diagram_recommendations(state)
    return state


def run_all_checks_step9(state: DocumentState) -> DocumentState:
    """Step 9: adds generated diagram checks."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    check_accuracy_findings(state)
    check_missing_quizzes(state)
    check_diagram_recommendations(state)
    check_generated_diagrams(state)
    return state


def run_all_checks_step10(state: DocumentState) -> DocumentState:
    """Step 10: adds table and code findings."""
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    check_accuracy_findings(state)
    check_missing_quizzes(state)
    check_diagram_recommendations(state)
    check_generated_diagrams(state)
    check_table_findings(state)
    check_code_findings(state)
    return state


def run_all_checks_step11(state: DocumentState) -> DocumentState:
    check_heading_styles(state)
    check_body_styles(state)
    check_heading_hierarchy(state)
    check_navigation_compatibility(state)
    check_spacing_issues(state)
    check_grammar_spelling(state)
    check_image_presence(state)
    check_ocr_results(state)
    check_image_caption_and_source(state)
    check_duplicates(state)
    check_retrieval_similarity(state)
    check_accuracy_findings(state)
    check_missing_quizzes(state)
    check_diagram_recommendations(state)
    check_visual_specs(state)
    check_generated_visuals(state)
    check_table_findings(state)
    check_code_findings(state)
    return state