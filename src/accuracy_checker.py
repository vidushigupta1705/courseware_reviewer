import json
import os
from typing import List, Dict

from mistralai.client import Mistral

from config import (
    ENABLE_ACCURACY_CHECK,
    ACCURACY_MODEL,
    MAX_UNITS_FOR_ACCURACY,
    ACCURACY_MIN_UNIT_TEXT_CHARS,
    MAX_ACCURACY_FINDINGS_PER_UNIT,
)
from models import AccuracyFinding


def _build_accuracy_prompt(unit_title: str, unit_text: str) -> str:
    return f"""
You are a technical courseware reviewer checking content accuracy.

Review the following unit content for factual inaccuracies, misleading statements,
outdated information, or unclear/ambiguous claims that could confuse learners.

Rules:
1. Only flag genuine factual problems — do not flag style or grammar issues.
2. Do not flag things that are merely simplified explanations appropriate for courseware.
3. For each issue found, identify the specific sentence or phrase that is problematic.
4. If no inaccuracies are found, return an empty list.
5. Return strict JSON only with one key: findings
6. findings must be a list of objects, each with:
   - flagged_text: the exact sentence or phrase that is inaccurate or misleading (keep short, under 200 chars)
   - issue_description: what is wrong with it (1-2 sentences)
   - suggestion: how it should be corrected (1-2 sentences)

Unit title:
{unit_title}

Unit content:
{unit_text}
""".strip()


def _call_mistral_accuracy(client: Mistral, prompt: str) -> List[Dict]:
    response = client.chat.complete(
        model=ACCURACY_MODEL,
        messages=[
            {"role": "system", "content": "You produce strict JSON only."},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.1,
        random_seed=42,
    )

    content = response.choices[0].message.content
    if isinstance(content, list):
        content = "".join(
            chunk.get("text", "") if isinstance(chunk, dict) else str(chunk)
            for chunk in content
        )

    parsed = json.loads(content)
    findings = parsed.get("findings", [])
    if not isinstance(findings, list):
        return []
    return findings


def _find_paragraph_index_for_text(state, unit, flagged_text: str) -> int:
    """
    Try to find the paragraph index in the unit that contains the flagged text.
    Falls back to the first non-heading body paragraph in the unit.
    Never falls back to a heading paragraph — headings are skipped by the rewrite pipeline.
    """
    search = flagged_text.strip().lower()[:80]
    for idx in range(unit.start_paragraph_index, unit.end_paragraph_index + 1):
        if idx < 0 or idx >= len(state.paragraphs):
            continue
        para_text = (state.paragraphs[idx].text or "").lower()
        if search and search[:40] in para_text:
            return idx

    # Fallback: first non-heading paragraph with enough text in the unit
    for idx in range(unit.start_paragraph_index + 1, unit.end_paragraph_index + 1):
        if idx < 0 or idx >= len(state.paragraphs):
            continue
        para = state.paragraphs[idx]
        if not para.is_heading and len((para.text or "").strip()) >= 60:
            return idx

    # Last resort: start_paragraph_index + 1
    return min(unit.start_paragraph_index + 1, unit.end_paragraph_index)


def run_accuracy_check(state) -> object:
    """
    For each detected unit, ask Mistral to review the content for factual
    inaccuracies. Results are stored in state.accuracy_findings and will
    be converted to ReviewIssue comments by check_accuracy_findings().
    """
    if not ENABLE_ACCURACY_CHECK:
        return state

    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        return state

    if not getattr(state, "units", []):
        return state

    client = Mistral(api_key=api_key)
    all_findings = []

    units_to_check = [
        u for u in state.units
        if u.body_text and len(u.body_text.strip()) >= ACCURACY_MIN_UNIT_TEXT_CHARS
    ][:MAX_UNITS_FOR_ACCURACY]

    for unit in units_to_check:
        try:
            prompt = _build_accuracy_prompt(unit.title, unit.body_text[:10000])
            raw_findings = _call_mistral_accuracy(client, prompt)

            unit_findings = []
            for item in raw_findings:
                flagged_text = (item.get("flagged_text") or "").strip()
                issue_description = (item.get("issue_description") or "").strip()
                suggestion = (item.get("suggestion") or "").strip()

                if not flagged_text or not issue_description:
                    continue

                para_idx = _find_paragraph_index_for_text(state, unit, flagged_text)

                unit_findings.append(
                    AccuracyFinding(
                        unit_title=unit.title,
                        paragraph_index=para_idx,
                        issue_description=issue_description,
                        flagged_text=flagged_text,
                        suggestion=suggestion,
                    )
                )

            # Cap findings per unit to avoid overwhelming comments
            all_findings.extend(unit_findings[:MAX_ACCURACY_FINDINGS_PER_UNIT])

        except Exception:
            continue

    state.accuracy_findings = all_findings
    return state