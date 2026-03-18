import re
from collections import Counter, defaultdict

from config import QUIZ_MIN_UNIT_TEXT_CHARS


CHAPTER_KEYWORDS = ["chapter", "unit", "module", "lesson"]
TOPIC_KEYWORDS = [
    "topic", "subtopic", "overview", "introduction", "summary",
    "example", "examples", "exercise", "practice", "review"
]

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


def _matches_numbered_chapter_pattern(text: str) -> bool:
    """
    Matches things like:
    - Chapter 1
    - Unit 2
    - Module 3
    - Lesson 4
    - 1 Introduction
    - 2. Networking Basics
    - 3: Security
    """
    text = _norm(text)

    patterns = [
        r"^(chapter|unit|module|lesson)\s+\d+\b",
        r"^(chapter|unit|module|lesson)\s+[ivx]+\b",
        r"^\d+\s+[A-Za-z].+",
        r"^\d+[.:]\s*[A-Za-z].+",
        r"^\d+\.\d+\s+[A-Za-z].+",   # keep as possible, but score lower later
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

    # subsection numbering like 1.1, 2.4, 3.2.1
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
        })

    return candidates


def _estimate_body_lengths_between_candidates(candidates, total_paragraphs):
    """
    Compute paragraph span length until next candidate.
    """
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

    Priority:
    1. Heading level with strongest chapter keyword evidence
    2. Heading level with strongest numbered chapter evidence
    3. Heading level with repeated headings and substantial average body length
    4. Smallest heading level number among viable candidates
    """
    if not candidates:
        return None

    by_level = defaultdict(list)
    for item in candidates:
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

    # sort by best score, then prefer smaller heading level
    scored_levels.sort(key=lambda x: (-x["score"], x["level"]))

    # conservative viability checks
    for item in scored_levels:
        if item["count"] >= 2 and item["avg_span"] >= 3:
            return item["level"]

    return scored_levels[0]["level"] if scored_levels else None


def _filter_real_chapter_headings(candidates, chapter_level):
    """
    Keep only headings at chosen chapter level, then prefer:
    - chapter keywords
    - chapter numbering
    - not topic-like
    If no keyword headings exist, fallback to all non-topic headings at that level.
    """
    level_items = [x for x in candidates if x["heading_level"] == chapter_level]
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


def _chapter_has_quiz(paragraphs, start_idx, end_idx):
    """
    Check only inside this chapter, and only in heading paragraphs.
    Avoids false positives from regular body text that mentions
    words like 'questions' or 'exercise'.
    """
    for idx in range(start_idx + 1, end_idx + 1):
        para = paragraphs[idx]
        if not para.is_heading:
            continue
        text = para.text
        if text and _looks_like_quiz_text(text):
            return True
    return False


def detect_units(state):
    """
    Detect full chapter/unit boundaries only.
    Generate quiz once per chapter, never per topic.

    This function is intentionally conservative and robust across:
    - Chapter 1 / Unit 1 / Module 1 style headings
    - numbered chapter headings like '1 Introduction'
    - documents with inconsistent heading levels
    """
    candidates = _heading_candidates(state)
    if not candidates:
        return []

    chapter_level = _choose_chapter_heading_level(state, candidates)
    if chapter_level is None:
        return []

    chapter_headings = _filter_real_chapter_headings(candidates, chapter_level)
    chapter_headings = sorted(chapter_headings, key=lambda x: x["index"])

    if not chapter_headings:
        return []

    units = []

    for i, chapter in enumerate(chapter_headings):
        start_idx = chapter["index"]

        if i < len(chapter_headings) - 1:
            end_idx = chapter_headings[i + 1]["index"] - 1
        else:
            end_idx = len(state.paragraphs) - 1

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

    return units