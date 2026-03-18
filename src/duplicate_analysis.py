from collections import defaultdict
from rapidfuzz.fuzz import ratio

from utils import normalize_text


MIN_PARAGRAPH_LEN_FOR_DUP_CHECK = 25
NEAR_DUPLICATE_THRESHOLD = 92


def _normalized_for_dup(text: str) -> str:
    if not text:
        return ""
    return normalize_text(text).lower()


def analyze_duplicates(state):
    duplicate_findings = []

    paragraphs = state.paragraphs

    norm_to_indexes = defaultdict(list)
    for para in paragraphs:
        norm = _normalized_for_dup(para.text)
        if len(norm) >= MIN_PARAGRAPH_LEN_FOR_DUP_CHECK:
            norm_to_indexes[norm].append(para.index)

    for norm_text, idxs in norm_to_indexes.items():
        if len(idxs) > 1:
            keep_idx = idxs[0]
            for dup_idx in idxs[1:]:
                duplicate_findings.append({
                    "type": "exact_duplicate_paragraph",
                    "keep_index": keep_idx,
                    "duplicate_index": dup_idx,
                    "text_preview": paragraphs[dup_idx].text[:200],
                })

    heading_map = defaultdict(list)
    for para in paragraphs:
        if para.is_heading and para.text:
            key = _normalized_for_dup(para.text)
            if key:
                heading_map[key].append(para.index)

    for heading_text, idxs in heading_map.items():
        if len(idxs) > 1:
            keep_idx = idxs[0]
            for dup_idx in idxs[1:]:
                duplicate_findings.append({
                    "type": "repeated_heading",
                    "keep_index": keep_idx,
                    "duplicate_index": dup_idx,
                    "text_preview": paragraphs[dup_idx].text[:200],
                })

    candidate_indexes = [
        p.index for p in paragraphs
        if p.text and len(_normalized_for_dup(p.text)) >= MIN_PARAGRAPH_LEN_FOR_DUP_CHECK and not p.is_heading
    ]

    already_marked = set(
        (f["keep_index"], f["duplicate_index"])
        for f in duplicate_findings
        if f["type"] == "exact_duplicate_paragraph"
    )

    for i in range(len(candidate_indexes)):
        idx_i = candidate_indexes[i]
        text_i = _normalized_for_dup(paragraphs[idx_i].text)

        for j in range(i + 1, len(candidate_indexes)):
            idx_j = candidate_indexes[j]
            if (idx_i, idx_j) in already_marked:
                continue

            text_j = _normalized_for_dup(paragraphs[idx_j].text)

            if abs(len(text_i) - len(text_j)) > max(20, int(0.25 * max(len(text_i), len(text_j)))):
                continue

            sim = ratio(text_i, text_j)
            if sim >= NEAR_DUPLICATE_THRESHOLD:
                duplicate_findings.append({
                    "type": "near_duplicate_paragraph",
                    "keep_index": idx_i,
                    "duplicate_index": idx_j,
                    "similarity": sim,
                    "text_preview": paragraphs[idx_j].text[:200],
                })

    repeated_heading_groups = defaultdict(list)
    for finding in duplicate_findings:
        if finding["type"] == "repeated_heading":
            repeated_heading_groups[finding["keep_index"]].append(finding["duplicate_index"])

    for keep_idx, dup_idxs in repeated_heading_groups.items():
        duplicate_findings.append({
            "type": "possible_duplicate_topic",
            "keep_index": keep_idx,
            "duplicate_indexes": dup_idxs,
            "text_preview": paragraphs[keep_idx].text[:200],
        })

    state.duplicate_findings = duplicate_findings
    return state
