import json
import os
from typing import List, Dict, Set
from concurrent.futures import ThreadPoolExecutor, as_completed

from mistralai.client import Mistral

from config import (
    ENABLE_LLM_REWRITE,
    MISTRAL_REWRITE_MODEL,
    MAX_REWRITE_PARAGRAPHS,
    REWRITE_MIN_PARAGRAPH_CHARS,
    ENABLE_OLLAMA_FALLBACK,
    OLLAMA_REWRITE_MODEL,
)
from models import RewriteSuggestion
from ollama_client import call_ollama_json


# ───────────────────────────────────────────────
# Constants
# ───────────────────────────────────────────────

REWRITE_ELIGIBLE_ISSUES = {
    "possible_open_source_similarity",
    "spacing",
    "content_accuracy",
    "formatting_body",
    "grammar",
    "spelling",
}

IBM_SKIP_ISSUES = {
    "possible_ibm_similarity",
}


# ───────────────────────────────────────────────
# Target Collection
# ───────────────────────────────────────────────

def _collect_rewrite_targets(state) -> List[Dict]:
    issue_map: Dict[int, Set[str]] = {}
    accuracy_fixes: Dict[int, str] = {}

    for issue in state.issues:
        idx = issue.paragraph_index

        if idx is None or idx < 0 or idx >= len(state.paragraphs):
            continue

        issue_map.setdefault(idx, set()).add(issue.issue_type)

        if issue.issue_type == "content_accuracy" and issue.suggested_fix:
            accuracy_fixes[idx] = issue.suggested_fix

    targets = []

    for para_idx, issue_types in issue_map.items():
        para = state.paragraphs[para_idx]
        text = (para.text or "").strip()

        if para.is_heading:
            continue
        if getattr(para, "is_code_like", False):
            continue
        if len(text) < REWRITE_MIN_PARAGRAPH_CHARS:
            continue

        if issue_types & IBM_SKIP_ISSUES:
            targets.append({
                "paragraph_index": para_idx,
                "reason": "ibm_similarity_skip",
                "text": text,
                "skip": True,
                "skip_reason": (
                    "IBM-related similarity is flagged for review "
                    "but not auto-rewritten."
                ),
                "suggested_fix": None,
            })
            continue

        eligible = issue_types & REWRITE_ELIGIBLE_ISSUES
        if eligible:
            targets.append({
                "paragraph_index": para_idx,
                "reason": ",".join(sorted(eligible)),
                "text": text,
                "skip": False,
                "skip_reason": None,
                "suggested_fix": accuracy_fixes.get(para_idx),
            })

    targets.sort(key=lambda x: x["paragraph_index"])
    return targets[:MAX_REWRITE_PARAGRAPHS]


# ───────────────────────────────────────────────
# Prompt Builder
# ───────────────────────────────────────────────

def _build_rewrite_prompt(text: str, reason: str, suggested_fix: str = None) -> str:
    fix_instruction = ""

    if suggested_fix and "content_accuracy" in reason:
        fix_instruction = f"\nSpecific correction required:\n{suggested_fix}\n"

    return f"""
You are rewriting courseware content.

Task:
Rewrite the paragraph below to fix all issues identified in the reason.

Always apply ALL of the following improvements:
1. Fix grammar errors
2. Fix spelling mistakes
3. Improve punctuation and sentence structure
4. Remove redundancy and extra spacing
5. Fix factual issues if content_accuracy is present
6. Ensure originality if open_source_similarity is present
7. Preserve meaning and technical accuracy
8. Do NOT introduce new information

Return STRICT JSON:
{{
  "rewritten_text": "...",
  "summary_reason": "..."
}}

{fix_instruction}

Reason:
{reason}

Paragraph:
{text}
""".strip()


# ───────────────────────────────────────────────
# LLM Call
# ───────────────────────────────────────────────

def _call_mistral_json(client: Mistral, prompt: str) -> Dict:
    response = client.chat.complete(
        model=MISTRAL_REWRITE_MODEL,
        messages=[
            {"role": "system", "content": "You produce strict JSON only."},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.1,
        random_seed=7,
    )

    content = response.choices[0].message.content

    if isinstance(content, list):
        content = "".join(
            chunk.get("text", "") if isinstance(chunk, dict) else str(chunk)
            for chunk in content
        )

    return json.loads(content)


# ───────────────────────────────────────────────
# Worker Function (Parallel)
# ───────────────────────────────────────────────

def _rewrite_single(client, use_ollama, target):
    para_idx = target["paragraph_index"]
    original_text = target["text"]
    reason = target["reason"]

    if target["skip"]:
        return RewriteSuggestion(
            paragraph_index=para_idx,
            reason=reason,
            original_text=original_text,
            rewritten_text=original_text,
            applied=False,
            skipped=True,
            skip_reason=target["skip_reason"],
        )

    try:
        prompt = _build_rewrite_prompt(
            original_text,
            reason,
            target.get("suggested_fix")
        )

        if use_ollama:
            result = call_ollama_json(OLLAMA_REWRITE_MODEL, "You produce strict JSON only.", prompt)
        else:
            result = _call_mistral_json(client, prompt)

        rewritten_text = (result.get("rewritten_text") or "").strip()

        if not rewritten_text:
            rewritten_text = original_text

        return RewriteSuggestion(
            paragraph_index=para_idx,
            reason=reason,
            original_text=original_text,
            rewritten_text=rewritten_text,
            applied=(rewritten_text != original_text),
            skipped=False,
            skip_reason=None,
        )

    except Exception as e:
        return RewriteSuggestion(
            paragraph_index=para_idx,
            reason=reason,
            original_text=original_text,
            rewritten_text=original_text,
            applied=False,
            skipped=True,
            skip_reason=f"Rewrite failed: {str(e)}",
        )


# ───────────────────────────────────────────────
# Main Function
# ───────────────────────────────────────────────

def run_llm_rewrite(state):
    if not ENABLE_LLM_REWRITE:
        return state

    mistral_key = os.environ.get("MISTRAL_API_KEY")
    use_ollama  = not mistral_key and ENABLE_OLLAMA_FALLBACK

    if not mistral_key and not use_ollama:
        return state

    client = Mistral(api_key=mistral_key) if mistral_key else None
    targets = _collect_rewrite_targets(state)

    suggestions: List[RewriteSuggestion] = []

    MAX_WORKERS = 5

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(_rewrite_single, client, use_ollama, t) for t in targets]

        for future in as_completed(futures):
            suggestions.append(future.result())

    suggestions.sort(key=lambda x: x.paragraph_index)

    state.rewrite_suggestions = suggestions
    return state
