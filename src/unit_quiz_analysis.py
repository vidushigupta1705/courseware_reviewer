import re
from collections import Counter, defaultdict

from config import QUIZ_MIN_UNIT_TEXT_CHARS


CHAPTER_KEYWORDS = ["chapter", "unit", "module", "lesson"]

TOPIC_KEYWORDS = [
    "topic", "subtopic", "overview", "introduction", "summary",
    "example", "examples", "exercise", "practice", "review",
    "agenda", "objectives", "objective", "prerequisite", "prerequisites",
    "about this", "what you will", "what you'll", "at a glance",
    "table of contents", "contents", "scope", "abstract"
]

# Headings that look like structural nav elements, never real chapter units.
# A heading matching any of these should never anchor a quiz.
_SKIP_AS_UNIT = {
    "agenda", "table of contents", "contents", "toc",
    "objectives", "learning objectives", "learning outcomes",
    "prerequisites", "pre-requisites", "scope", "abstract",
    "about this course", "about this module", "what you will learn",
    "what you'll learn", "at a glance", "overview",
}

# Back-matter section titles (after stripping a trailing "— Unit N" /
# "- Chapter N" suffix) that should never anchor a quiz unit, even though
# they contain a chapter keyword like "Unit" in that suffix.
_BACK_MATTER_KEYWORDS = {
    "glossary", "glossary of key terms",
    "references", "references and further reading",
    "further reading", "bibliography", "appendix",
    "appendices", "index",
}

QUIZ_PATTERNS = [
    r"\bquiz\b",
    r"\bmcq\b",
    r"\bmultiple choice questions\b",
    r"\btrue\s*/\s*false\b",
    r"\btrue or false\b",
    r"\bfill in the blanks\b",
    r"\bself[- ]assessment\b",
    r"\breview questions\b",
    r"\bpractice questions\b",
    r"\bcheck your understanding\b",
    # "Assessment" as a standalone section title (e.g. "Assessment",
    # "Unit Assessment", "Assessment Questions:") — anchored to the WHOLE
    # line, not a bare \bassessment\b substring. A bare substring would
    # also match incidental uses like a title-page tagline ("Assessment
    # Included") or a sentence mentioning assessment in passing, which
    # are not quiz sections and must not trigger a false positive.
    r"^\s*(unit\s+|chapter\s+|module\s+|self[- ])?assessment\s*[:\-]?\s*(questions?)?\s*$",
]


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def _looks_like_quiz_text(text: str) -> bool:
    if not text:
        return False
    lower = text.lower()
    return any(re.search(pattern, lower, flags=re.IGNORECASE) for pattern in QUIZ_PATTERNS)


def _has_chapter_keyword(text: str) -> bool:
    lower = (text or "").lower()
    return any(word in lower for word in CHAPTER_KEYWORDS)


def _has_topic_keyword(text: str) -> bool:
    lower = (text or "").lower()
    return any(word in lower for word in TOPIC_KEYWORDS)


def _is_nav_heading(text: str) -> bool:
    """
    Returns True if this heading is a structural navigation element
    (agenda, TOC, objectives, etc.) or back-matter (glossary, references)
    that should never anchor a quiz unit.
    Strips leading numbering before matching so '1. Agenda' is caught too.
    """
    norm = _norm(text).lower()
    norm = re.sub(r"^\d+[\.\d]*\s*", "", norm).strip()
    if norm in _SKIP_AS_UNIT:
        return True
    # Back-matter headings often end with "— Unit N" / "- Unit N" which
    # would otherwise be mistaken for a real chapter/unit boundary because
    # they contain a chapter keyword. Check independent of trailing
    # "Unit N" / "Chapter N" suffixes.
    stripped = re.sub(r"[\u2013\u2014\-:]\s*(unit|chapter|module|lesson)\s+\d+\s*$", "", norm).strip()
    return stripped in _BACK_MATTER_KEYWORDS


def _matches_numbered_chapter_pattern(text: str) -> bool:
    """
    Matches things like:
    - Chapter 1 / Unit 2 / Module 3 / Lesson 4
    - 1 Introduction
    - 2. Networking Basics
    - 3: Security
    """
    text = _norm(text)
    patterns = [
        r"^(chapter|unit|module|lesson)\s+\d+\b",
        r"^(chapter|unit|module|lesson)\s+[ivx]+\b",
        r"^\d+\s+[A-Za-z].+",
        r"^\d+[.:]\s+[A-Za-z].+",
    ]
    return any(re.search(p, text, flags=re.IGNORECASE) for p in patterns)


def _is_likely_topic_heading(text: str) -> bool:
    """
    Topic-like headings often:
    - contain topic/subtopic words
    - look like 1.1 / 2.3 style subsections
    """
    text = _norm(text)
    lower = text.lower()

    if _has_topic_keyword(lower):
        return True

    if re.match(r"^\d+\.\d+(\.\d+)*\b", text):
        return True

    return False


def _heading_candidates(state):
    candidates = []

    for para in state.paragraphs:
        if not para.is_heading:
            continue
        if not para.text or not para.text.strip():
            continue

        text = _norm(para.text)

        candidates.append({
            "index": para.index,
            "text": text,
            "heading_level": para.heading_level,
            "has_chapter_keyword": _has_chapter_keyword(text),
            "matches_numbered_pattern": _matches_numbered_chapter_pattern(text),
            "is_likely_topic": _is_likely_topic_heading(text),
            "is_nav_heading": _is_nav_heading(text),       # new field
        })

    return candidates


def _estimate_body_lengths_between_candidates(candidates, total_paragraphs):
    spans = []
    for i, item in enumerate(candidates):
        start_idx = item["index"]
        if i < len(candidates) - 1:
            end_idx = candidates[i + 1]["index"] - 1
        else:
            end_idx = total_paragraphs - 1
        spans.append(max(0, end_idx - start_idx))
    return spans


def _choose_chapter_heading_level(state, candidates):
    """
    Strict, deterministic chapter-level selection.
    Nav headings are excluded before scoring so they can't skew the result.
    """
    if not candidates:
        return None

    # Nav headings are never chapter anchors — exclude from level scoring.
    scoreable = [c for c in candidates if not c["is_nav_heading"]]
    if not scoreable:
        return None

    by_level = defaultdict(list)
    for item in scoreable:
        by_level[item["heading_level"]].append(item)

    scored_levels = []

    for level, items in by_level.items():
        keyword_count = sum(1 for x in items if x["has_chapter_keyword"])
        numbered_count = sum(1 for x in items if x["matches_numbered_pattern"])
        topic_like_count = sum(1 for x in items if x["is_likely_topic"])

        items_sorted = sorted(items, key=lambda x: x["index"])
        spans = _estimate_body_lengths_between_candidates(items_sorted, len(state.paragraphs))
        avg_span = sum(spans) / len(spans) if spans else 0

        score = 0
        score += keyword_count * 10
        score += numbered_count * 6
        score += len(items) * 2
        score += int(avg_span)
        score -= topic_like_count * 8

        scored_levels.append({
            "level": level,
            "score": score,
            "keyword_count": keyword_count,
            "numbered_count": numbered_count,
            "count": len(items),
            "avg_span": avg_span,
            "topic_like_count": topic_like_count,
        })

    scored_levels.sort(key=lambda x: (-x["score"], x["level"]))

    for item in scored_levels:
        if item["count"] >= 2 and item["avg_span"] >= 3:
            return item["level"]

    return scored_levels[0]["level"] if scored_levels else None


def _filter_real_chapter_headings(candidates, chapter_level):
    """
    Keep only headings at the chosen chapter level.
    Nav headings are always excluded regardless of level match.
    Falls back progressively: strong → non-topic → all at that level.
    """
    level_items = [
        x for x in candidates
        if x["heading_level"] == chapter_level and not x["is_nav_heading"]   # nav always excluded
    ]
    if not level_items:
        return []

    strong = [
        x for x in level_items
        if (x["has_chapter_keyword"] or x["matches_numbered_pattern"]) and not x["is_likely_topic"]
    ]
    if strong:
        return strong

    non_topic = [x for x in level_items if not x["is_likely_topic"]]
    if non_topic:
        return non_topic

    return level_items


def _is_quiz_pseudo_heading(para) -> bool:
    """
    Return True for short paragraphs that act as quiz section titles
    even when they carry no Word heading style (e.g. 'Quiz', 'MCQ', 'Review Questions'
    authored as bold Normal-style paragraphs).
    """
    text = (para.text or "").strip()
    if not text or len(text.split()) > 12:
        return False
    return _looks_like_quiz_text(text)


def _chapter_has_quiz(paragraphs, start_idx, end_idx):
    """
    Detect whether this chapter already contains a quiz section.

    Checks:
    1. Word-styled heading paragraphs with quiz keywords.
    2. Short pseudo-heading paragraphs with quiz keywords (bold Normal-style).
    3. MCQ body structure: numbered question followed by 2+ lettered options.
    """
    mcq_option_count = 0
    MCQ_OPTION_PATTERN = re.compile(r'^\s*[A-Da-d][)\.]\s*\S')
    MCQ_QUESTION_PATTERN = re.compile(r'^\s*\d+[\.\)]\s*\S')
    # Inline options pattern: "A) ... B) ... C)" all on one line
    INLINE_OPTIONS_PATTERN = re.compile(
        r'[A-D][)\.].+[A-D][)\.].+[A-D][\.\)]', re.IGNORECASE
    )
    in_numbered_question = False

    for idx in range(start_idx + 1, end_idx + 1):
        para = paragraphs[idx]
        text = para.text or ""

        # Signal 1 & 2: heading or short pseudo-heading with quiz keyword
        if _is_quiz_pseudo_heading(para):
            return True
        # Signal 4: inline MCQ — question with options all on the same line
        if INLINE_OPTIONS_PATTERN.search(text):
            return True

        # Signal 3: MCQ structure — numbered question + 2 lettered options
        if MCQ_QUESTION_PATTERN.match(text):
            in_numbered_question = True
            mcq_option_count = 0
        elif MCQ_OPTION_PATTERN.match(text) and in_numbered_question:
            mcq_option_count += 1
            if mcq_option_count >= 2:
                return True
        else:
            if text.strip():
                in_numbered_question = False
                mcq_option_count = 0

    return False


def _build_whole_document_unit(state) -> list:
    """
    Fallback: treat the entire document as a single lesson unit.
    Used when no reliable chapter-level headings are found.
    The quiz will be appended at the very end of the document.
    """
    paragraphs = state.paragraphs
    if not paragraphs:
        return []

    body_parts = [p.text for p in paragraphs if p.text and p.text.strip()]
    body_text = "\n".join(body_parts).strip()

    if len(body_text) < QUIZ_MIN_UNIT_TEXT_CHARS:
        return []

    # Use the first heading as the unit title if one exists, else a generic label.
    title = next(
        (p.text.strip() for p in paragraphs if p.is_heading and p.text and p.text.strip()),
        "Lesson",
    )

    quiz_present = _chapter_has_quiz(
        paragraphs,
        start_idx=0,
        end_idx=len(paragraphs) - 1,
    )

    return [{
        "title": title,
        "heading_paragraph_index": paragraphs[0].index,
        "heading_level": 1,
        "start_paragraph_index": paragraphs[0].index,
        "end_paragraph_index": paragraphs[-1].index,   # always the true end
        "body_text": body_text,
        "quiz_present": quiz_present,
    }]


def detect_units(state):
    """
    Detect full chapter/unit boundaries only.
    One quiz per chapter, placed at the end of that chapter's content.

    Falls back to treating the whole document as one unit when no reliable
    chapter headings are found — this ensures the quiz always goes at the end,
    never near an agenda or TOC heading near the top.
    """
    candidates = _heading_candidates(state)
    if not candidates:
        return _build_whole_document_unit(state)

    chapter_level = _choose_chapter_heading_level(state, candidates)
    if chapter_level is None:
        return _build_whole_document_unit(state)

    chapter_headings = _filter_real_chapter_headings(candidates, chapter_level)
    chapter_headings = sorted(chapter_headings, key=lambda x: x["index"])

    if not chapter_headings:
        return _build_whole_document_unit(state)

    units = []

    for i, chapter in enumerate(chapter_headings):
        start_idx = chapter["index"]

        if i < len(chapter_headings) - 1:
            end_idx = chapter_headings[i + 1]["index"] - 1
        else:
            end_idx = len(state.paragraphs) - 1    # last chapter runs to true document end

        if end_idx <= start_idx:
            continue

        body_parts = []
        for idx in range(start_idx + 1, end_idx + 1):
            text = state.paragraphs[idx].text
            if text and text.strip():
                body_parts.append(text)

        body_text = "\n".join(body_parts).strip()
        if len(body_text) < QUIZ_MIN_UNIT_TEXT_CHARS:
            continue

        quiz_present = _chapter_has_quiz(state.paragraphs, start_idx, end_idx)

        units.append({
            "title": chapter["text"],
            "heading_paragraph_index": chapter["index"],
            "heading_level": chapter["heading_level"],
            "start_paragraph_index": start_idx,
            "end_paragraph_index": end_idx,
            "body_text": body_text,
            "quiz_present": quiz_present,
        })

    # If all detected chapters were filtered out (e.g. all too short),
    # fall back to whole-document unit so we don't silently produce nothing.
    return units if units else _build_whole_document_unit(state)
