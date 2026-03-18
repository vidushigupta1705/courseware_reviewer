import json
import os
from typing import List, Dict, Set

from mistralai.client import Mistral

from config import (
    ENABLE_LLM_REWRITE,
    MISTRAL_REWRITE_MODEL,
    MAX_REWRITE_PARAGRAPHS,
    REWRITE_MIN_PARAGRAPH_CHARS,
)
from models import RewriteSuggestion


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


def _collect_rewrite_targets(state) -> List[Dict]:
    issue_map: Dict[int, Set[str]] = {}
    # Also collect suggested_fix for content_accuracy issues
    accuracy_fixes: Dict[int, str] = {}

    for issue in state.issues:
        if issue.paragraph_index is None:
            continue
        if issue.paragraph_index < 0 or issue.paragraph_index >= len(state.paragraphs):
            continue
        issue_map.setdefault(issue.paragraph_index, set()).add(issue.issue_type)
        if issue.issue_type == "content_accuracy" and issue.suggested_fix:
            accuracy_fixes[issue.paragraph_index] = issue.suggested_fix

    targets = []
    for para_idx, issue_types in issue_map.items():
        para = state.paragraphs[para_idx]
        text = para.text or ""

        if para.is_heading:
            continue
        if getattr(para, "is_code_like", False):
            continue
        if len(text.strip()) < REWRITE_MIN_PARAGRAPH_CHARS:
            continue

        if issue_types & IBM_SKIP_ISSUES:
            targets.append({
                "paragraph_index": para_idx,
                "reason": "ibm_similarity_skip",
                "text": text,
                "skip": True,
                "skip_reason": "IBM-related similarity is flagged for review but not auto-rewritten in Step 6+.",
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

    targets = sorted(targets, key=lambda x: x["paragraph_index"])
    return targets[:MAX_REWRITE_PARAGRAPHS]

def _build_rewrite_prompt(text: str, reason: str, suggested_fix: str = None) -> str:
    fix_instruction = ""
    if suggested_fix and "content_accuracy" in reason:
        fix_instruction = f"\nSpecific correction required:\n{suggested_fix}\n"

    return f"""
You are rewriting courseware content.

Task:
Rewrite the paragraph below to fix all issues identified in the reason.
Always apply ALL of the following improvements regardless of reason:
1. Fix all grammar errors.
2. Fix all spelling mistakes.
3. Fix punctuation and sentence structure.
4. Remove unnecessary spacing or redundant wording.
5. If the reason mentions content_accuracy, correct the factual inaccuracy described.
6. If the reason mentions open_source_similarity, rephrase to ensure originality.
7. Preserve all technical terms, meaning, and courseware tone.
8. Do not add new facts.
9. Return strict JSON only with keys:
   - rewritten_text
   - summary_reason
{fix_instruction}
Reason:
{reason}

Paragraph:
{text}
""".strip()


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

    parsed = json.loads(content)
    return parsed


def run_llm_rewrite(state):
    if not ENABLE_LLM_REWRITE:
        return state

    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        return state

    client = Mistral(api_key=api_key)
    targets = _collect_rewrite_targets(state)

    suggestions = []

    for target in targets:
        para_idx = target["paragraph_index"]
        original_text = target["text"]
        reason = target["reason"]

        if target["skip"]:
            suggestions.append(
                RewriteSuggestion(
                    paragraph_index=para_idx,
                    reason=reason,
                    original_text=original_text,
                    rewritten_text=original_text,
                    applied=False,
                    skipped=True,
                    skip_reason=target["skip_reason"],
                )
            )
            continue

        try:
            prompt = _build_rewrite_prompt(original_text, reason, target.get("suggested_fix"))
            result = _call_mistral_json(client, prompt)
            rewritten_text = (result.get("rewritten_text") or "").strip()

            if not rewritten_text:
                rewritten_text = original_text

            suggestions.append(
                RewriteSuggestion(
                    paragraph_index=para_idx,
                    reason=reason,
                    original_text=original_text,
                    rewritten_text=rewritten_text,
                    applied=True if rewritten_text != original_text else False,
                    skipped=False,
                    skip_reason=None,
                )
            )
        except Exception as e:
            suggestions.append(
                RewriteSuggestion(
                    paragraph_index=para_idx,
                    reason=reason,
                    original_text=original_text,
                    rewritten_text=original_text,
                    applied=False,
                    skipped=True,
                    skip_reason=f"Rewrite failed: {str(e)}",
                )
            )

    state.rewrite_suggestions = suggestions
    return state