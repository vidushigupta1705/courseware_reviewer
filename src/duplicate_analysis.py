import re
import logging
from collections import defaultdict
from rapidfuzz.fuzz import ratio
from sentence_transformers import SentenceTransformer
import numpy as np

from utils import normalize_text
from config import (
    MIN_PARAGRAPH_LEN_FOR_DUP_CHECK,
    NEAR_DUPLICATE_THRESHOLD,
    JACCARD_SIMILARITY_THRESHOLD,
    SEMANTIC_SIMILARITY_THRESHOLD,
    MAX_NEAR_DUP_PER_PARAGRAPH,
    ENABLE_SEMANTIC_DUPLICATES,
    EMBEDDING_MODEL_NAME,
)

logger = logging.getLogger(__name__)

# If a paragraph repeats more than this many times, it is a structural/template
# element (page header, course title, boilerplate). Flag it ONCE with a count.
STRUCTURAL_REPEAT_THRESHOLD = 5

STRUCTURAL_HEADINGS = {
    "summary", "overview", "introduction", "objectives",
    "learning objectives", "learning outcomes", "conclusion", "conclusions",
    "review", "review questions", "exercises", "exercise", "practice",
    "references", "reference", "further reading", "key takeaways",
    "key concepts", "key terms", "recap", "what you learned",
    "what we covered", "check your understanding", "self-assessment",
    "self assessment", "quiz", "pre-requisites", "prerequisites",
    "agenda", "topics covered", "table of contents", "contents",
    "abstract", "note", "notes", "tip", "tips", "important",
    "warning", "caution",
}

_embedding_model = None

def _get_model():
    global _embedding_model
    if _embedding_model is None:
        logger.info(
            f"[Semantic Duplicates] Loading sentence-transformer model "
            f"'{EMBEDDING_MODEL_NAME}'. First run may download ~90MB."
        )
        _embedding_model = SentenceTransformer(EMBEDDING_MODEL_NAME)
        logger.info("[Semantic Duplicates] Model loaded successfully.")
    return _embedding_model


def _normalized_for_dup(text: str) -> str:
    return normalize_text(text).lower() if text else ""


def _jaccard_fast(a: str, b: str) -> float:
    set_a, set_b = set(a.split()), set(b.split())
    if not set_a or not set_b:
        return 0.0
    return len(set_a & set_b) / len(set_a | set_b)


def _is_structural_heading(text: str) -> bool:
    if not text:
        return False
    norm = normalize_text(text).lower().strip()
    norm = re.sub(r"^\d+[\.\d]*\s*", "", norm).strip()
    return norm in STRUCTURAL_HEADINGS


def analyze_duplicates(state):
    paragraphs = state.paragraphs
    duplicate_findings = []

    norm_cache = {
        p.index: _normalized_for_dup(p.text)
        for p in paragraphs if p.text
    }

    # ── 1. Exact duplicates ───────────────────────────────────────────────────
    norm_map = defaultdict(list)
    for p in paragraphs:
        norm = norm_cache.get(p.index, "")
        if len(norm) >= MIN_PARAGRAPH_LEN_FOR_DUP_CHECK:
            norm_map[norm].append(p.index)

    structural_norms = set()

    for norm, idxs in norm_map.items():
        if len(idxs) <= 1:
            continue

        if len(idxs) > STRUCTURAL_REPEAT_THRESHOLD:
            # Repeating template element — flag exactly ONCE with count
            structural_norms.add(norm)
            duplicate_findings.append({
                "type": "exact_duplicate_paragraph",
                "keep_index": idxs[0],
                "duplicate_index": idxs[1],
                "occurrence_count": len(idxs),
                "is_structural": True,
                "text_preview": paragraphs[idxs[0]].text[:200],
            })
        else:
            # Genuine content duplication — flag each occurrence
            keep = idxs[0]
            for dup in idxs[1:]:
                duplicate_findings.append({
                    "type": "exact_duplicate_paragraph",
                    "keep_index": keep,
                    "duplicate_index": dup,
                    "occurrence_count": len(idxs),
                    "is_structural": False,
                    "text_preview": paragraphs[dup].text[:200],
                })

    # ── 2. Repeated headings ──────────────────────────────────────────────────
    heading_map = defaultdict(list)
    for p in paragraphs:
        if p.is_heading and p.text:
            heading_map[norm_cache.get(p.index, "")].append(p.index)

    for key, idxs in heading_map.items():
        if len(idxs) <= 1:
            continue
        original_text = paragraphs[idxs[0]].text
        if _is_structural_heading(original_text):
            continue
        if len(idxs) > STRUCTURAL_REPEAT_THRESHOLD:
            continue
        keep = idxs[0]
        for dup in idxs[1:]:
            duplicate_findings.append({
                "type": "repeated_heading",
                "keep_index": keep,
                "duplicate_index": dup,
                "text_preview": paragraphs[dup].text[:200],
            })

    already_marked = set(
        (f["keep_index"], f["duplicate_index"])
        for f in duplicate_findings
        if "duplicate_index" in f
    )

    # Candidates: non-heading, long enough, not a structural repeating element
    candidate_indexes = [
        p.index for p in paragraphs
        if p.text
        and len(norm_cache.get(p.index, "")) >= MIN_PARAGRAPH_LEN_FOR_DUP_CHECK
        and not p.is_heading
        and norm_cache.get(p.index, "") not in structural_norms
    ]

    # ── 3. Fuzzy near-duplicate (full document, no cap) ───────────────────────
    per_para_count = defaultdict(int)

    for i in range(len(candidate_indexes)):
        idx_i = candidate_indexes[i]

        if per_para_count[idx_i] >= MAX_NEAR_DUP_PER_PARAGRAPH:
            continue

        text_i = norm_cache[idx_i]

        for j in range(i + 1, len(candidate_indexes)):
            idx_j = candidate_indexes[j]

            if (idx_i, idx_j) in already_marked:
                continue

            text_j = norm_cache[idx_j]

            len_i, len_j = len(text_i), len(text_j)
            if abs(len_i - len_j) > max(20, int(0.25 * max(len_i, len_j))):
                continue

            if _jaccard_fast(text_i, text_j) < JACCARD_SIMILARITY_THRESHOLD:
                continue

            sim = ratio(text_i, text_j)

            if sim >= NEAR_DUPLICATE_THRESHOLD:
                duplicate_findings.append({
                    "type": "near_duplicate_paragraph",
                    "keep_index": idx_i,
                    "duplicate_index": idx_j,
                    "similarity": round(sim, 2),
                    "is_structural": False,
                    "text_preview": paragraphs[idx_j].text[:200],
                })
                per_para_count[idx_i] += 1
                already_marked.add((idx_i, idx_j))

    # ── 4. Semantic duplicates (full document, batched) ───────────────────────
    if ENABLE_SEMANTIC_DUPLICATES and candidate_indexes:
        try:
            model = _get_model()
            all_texts = [paragraphs[i].text for i in candidate_indexes]
            logger.info(f"[Semantic Duplicates] Encoding {len(all_texts)} paragraphs…")

            embeddings = model.encode(
                all_texts,
                batch_size=64,
                normalize_embeddings=True,
                show_progress_bar=False,
            )
            logger.info("[Semantic Duplicates] Encoding complete.")

            n = len(candidate_indexes)
            for i in range(n):
                idx_i = candidate_indexes[i]
                for j in range(i + 1, n):
                    idx_j = candidate_indexes[j]

                    if (idx_i, idx_j) in already_marked:
                        continue

                    score = float(np.dot(embeddings[i], embeddings[j]))

                    if score >= SEMANTIC_SIMILARITY_THRESHOLD:
                        duplicate_findings.append({
                            "type": "semantic_duplicate_paragraph",
                            "keep_index": idx_i,
                            "duplicate_index": idx_j,
                            "similarity": round(score, 4),
                            "is_structural": False,
                            "text_preview": paragraphs[idx_j].text[:200],
                        })
                        already_marked.add((idx_i, idx_j))

        except Exception as e:
            logger.warning(f"[Semantic Duplicates] Skipped due to error: {e}")

    # ── 5. Topic grouping ─────────────────────────────────────────────────────
    grouped = defaultdict(list)
    for f in duplicate_findings:
        if f["type"] == "repeated_heading":
            grouped[f["keep_index"]].append(f["duplicate_index"])

    for keep, dup_idxs in grouped.items():
        duplicate_findings.append({
            "type": "possible_duplicate_topic",
            "keep_index": keep,
            "duplicate_indexes": dup_idxs,
            "text_preview": paragraphs[keep].text[:200],
        })

    logger.info(f"[Duplicates] Complete — {len(duplicate_findings)} total findings.")

    state.duplicate_findings = duplicate_findings
    return state
