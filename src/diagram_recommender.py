from config import (
    ENABLE_DIAGRAM_RECOMMENDATION,
    DIAGRAM_MIN_PARAGRAPH_CHARS,
    DIAGRAM_MAX_RECOMMENDATIONS,
    DIAGRAM_PROCESS_KEYWORDS,
)


def _score_diagram_need(text: str) -> int:
    if not text:
        return 0

    lower = text.lower()
    score = 0

    if len(text) >= DIAGRAM_MIN_PARAGRAPH_CHARS:
        score += 1

    keyword_hits = 0
    for kw in DIAGRAM_PROCESS_KEYWORDS:
        if kw in lower:
            keyword_hits += 1
    score += keyword_hits

    # sequence indicators
    sequence_markers = [
        "first", "second", "third", "next", "then", "finally",
        "step 1", "step 2", "step 3",
        "1.", "2.", "3."
    ]
    for marker in sequence_markers:
        if marker in lower:
            score += 1

    # arrow / mapping style wording
    mapping_markers = ["->", "=>", "maps to", "flows to", "transfers to", "moves to"]
    for marker in mapping_markers:
        if marker in lower:
            score += 1

    return score


def analyze_diagram_recommendations(state):
    if not ENABLE_DIAGRAM_RECOMMENDATION:
        state.diagram_recommendations = []
        return state

    recommendations = []

    for para in state.paragraphs:
        if para.is_heading:
            continue
        if not para.text or len(para.text.strip()) < DIAGRAM_MIN_PARAGRAPH_CHARS:
            continue

        score = _score_diagram_need(para.text)
        if score >= 3:
            recommendations.append({
                "paragraph_index": para.index,
                "score": score,
                "text_preview": para.text[:300],
                "recommendation_type": "diagram_recommended",
            })

    recommendations = sorted(recommendations, key=lambda x: (-x["score"], x["paragraph_index"]))
    state.diagram_recommendations = recommendations[:DIAGRAM_MAX_RECOMMENDATIONS]
    return state