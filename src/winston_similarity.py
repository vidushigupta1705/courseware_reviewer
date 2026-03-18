import os
import time
import requests

from config import (
    WINSTON_API_URL,
    ENABLE_WINSTON,
    WINSTON_LANGUAGE,
    WINSTON_COUNTRY,
    WINSTON_MIN_TEXT_CHARS,
    WINSTON_MAX_TEXT_CHARS,
    WINSTON_MAX_PARAGRAPHS,
    WINSTON_EXCLUDED_SOURCES,
    WINSTON_SCORE_FLAG_THRESHOLD,
    WINSTON_HIGH_SCORE_THRESHOLD,
    WINSTON_ONLY_SUSPICIOUS,
    WINSTON_FALLBACK_SCAN_IF_FEW,
    WINSTON_MIN_CANDIDATES,
)
from utils import normalize_text


def _clean_text_for_scan(text: str) -> str:
    return normalize_text(text)


def _eligible_text(text: str) -> bool:
    if not text:
        return False
    cleaned = _clean_text_for_scan(text)
    return WINSTON_MIN_TEXT_CHARS <= len(cleaned) <= WINSTON_MAX_TEXT_CHARS


def _truncate_text(text: str) -> str:
    cleaned = _clean_text_for_scan(text)
    if len(cleaned) > WINSTON_MAX_TEXT_CHARS:
        return cleaned[:WINSTON_MAX_TEXT_CHARS]
    return cleaned


def _call_winston_plagiarism_api(text: str):
    api_key = os.environ.get("WINSTON_API_KEY")
    if not api_key:
        return {"ok": False, "error": "Missing WINSTON_API_KEY", "response": None}

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    payload = {
        "text": _truncate_text(text),
        "excluded_sources": WINSTON_EXCLUDED_SOURCES,
        "language": WINSTON_LANGUAGE,
        "country": WINSTON_COUNTRY,
    }

    try:
        resp = requests.post(WINSTON_API_URL, headers=headers, json=payload, timeout=90)
        if resp.status_code != 200:
            return {"ok": False, "error": f"HTTP {resp.status_code}: {resp.text[:500]}", "response": None}
        return {"ok": True, "error": None, "response": resp.json()}
    except Exception as e:
        return {"ok": False, "error": str(e), "response": None}


def _duplicate_indexes(state):
    indexes = set()
    for finding in getattr(state, "duplicate_findings", []):
        if finding.get("type") in {"exact_duplicate_paragraph", "repeated_heading"}:
            dup = finding.get("duplicate_index")
            if dup is not None:
                indexes.add(dup)
    return indexes


def _suspicious_issue_indexes(state):
    suspicious_types = {
        "spacing",
        "near_duplicate_paragraph",
        "possible_duplicate_topic",
        "image_contains_text",
        "ocr_low_confidence",
        "ocr_low_text",
    }
    indexes = set()
    for issue in getattr(state, "issues", []):
        if issue.issue_type in suspicious_types and issue.paragraph_index is not None:
            indexes.add(issue.paragraph_index)
    return indexes


def _build_candidate_indexes(state):
    dup_indexes = _duplicate_indexes(state)
    suspicious_indexes = _suspicious_issue_indexes(state)

    candidates = []
    fallback = []

    for para in state.paragraphs:
        if para.is_heading:
            continue
        if para.index in dup_indexes:
            continue
        if not _eligible_text(para.text or ""):
            continue

        if para.index in suspicious_indexes:
            candidates.append(para.index)
        else:
            fallback.append(para.index)

    if WINSTON_ONLY_SUSPICIOUS:
        if len(candidates) >= WINSTON_MIN_CANDIDATES:
            return candidates[:WINSTON_MAX_PARAGRAPHS]
        if WINSTON_FALLBACK_SCAN_IF_FEW:
            merged = candidates + fallback
            return merged[:WINSTON_MAX_PARAGRAPHS]
        return candidates[:WINSTON_MAX_PARAGRAPHS]

    return (candidates + fallback)[:WINSTON_MAX_PARAGRAPHS]


def analyze_winston_similarity(state):
    if not ENABLE_WINSTON:
        state.retrieval_findings = []
        return state

    findings = []
    candidate_indexes = _build_candidate_indexes(state)

    for para_idx in candidate_indexes:
        para = state.paragraphs[para_idx]
        text = para.text or ""

        result = _call_winston_plagiarism_api(text)
        time.sleep(0.4)

        if not result["ok"]:
            findings.append({
                "type": "winston_scan_error",
                "paragraph_index": para.index,
                "severity": "low",
                "error": result["error"],
                "text_preview": text[:300],
            })
            continue

        data = result["response"] or {}
        result_obj = data.get("result", {})
        sources = data.get("sources", [])

        score = result_obj.get("score", 0)
        source_count = result_obj.get("sourceCounts", 0)
        total_plagiarism_words = result_obj.get("totalPlagiarismWords", 0)

        best_source = None
        if sources:
            best_source = max(sources, key=lambda s: s.get("score", 0))

        if score >= WINSTON_SCORE_FLAG_THRESHOLD:
            findings.append({
                "type": "winston_plagiarism_flag",
                "paragraph_index": para.index,
                "severity": "high" if score >= WINSTON_HIGH_SCORE_THRESHOLD else "medium",
                "score": score,
                "source_count": source_count,
                "total_plagiarism_words": total_plagiarism_words,
                "best_source_title": best_source.get("title") if best_source else "",
                "best_source_url": best_source.get("url") if best_source else "",
                "best_source_score": best_source.get("score") if best_source else 0,
                "best_source_description": best_source.get("description") if best_source else "",
                "best_source_is_excluded": best_source.get("is_excluded") if best_source else False,
                "matched_sequences": best_source.get("plagiarismFound", [])[:5] if best_source else [],
                "text_preview": text[:300],
            })

    state.retrieval_findings = findings
    return state