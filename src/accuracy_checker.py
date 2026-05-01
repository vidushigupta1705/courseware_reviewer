import json
import os
from typing import List, Dict
from concurrent.futures import ThreadPoolExecutor, as_completed

from mistralai.client import Mistral

from config import (
    ENABLE_ACCURACY_CHECK,
    ACCURACY_MODEL,
    MAX_UNITS_FOR_ACCURACY,
    ACCURACY_MIN_UNIT_TEXT_CHARS,
    MAX_ACCURACY_FINDINGS_PER_UNIT,
    ENABLE_OLLAMA_FALLBACK,
    OLLAMA_ACCURACY_MODEL,
)
from models import AccuracyFinding
from ollama_client import call_ollama_json


# ───────────────────────────────────────────────
# Prompt Builder
# ───────────────────────────────────────────────

def _build_accuracy_prompt(unit_title: str, unit_text: str) -> str:
    return f"""
You are a technical courseware reviewer checking content accuracy.

Review the following unit content for factual inaccuracies, misleading statements,
outdated information, or unclear/ambiguous claims.

Rules:
1. Only flag factual issues (NOT grammar/style).
2. Ignore acceptable simplifications.
3. Identify exact problematic sentence.
4. Return empty list if no issues.
5. Return STRICT JSON with key: findings
6. Each finding must include:
   - flagged_text
   - issue_description
   - suggestion

Unit title:
{unit_title}

Unit content:
{unit_text}
""".strip()


# ───────────────────────────────────────────────
# LLM Call
# ───────────────────────────────────────────────

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

    return findings if isinstance(findings, list) else []


# ───────────────────────────────────────────────
# Paragraph Mapping
# ───────────────────────────────────────────────

def _find_paragraph_index_for_text(state, unit, flagged_text: str) -> int:
    search = flagged_text.strip().lower()[:80]

    for idx in range(unit.start_paragraph_index, unit.end_paragraph_index + 1):
        if idx < 0 or idx >= len(state.paragraphs):
            continue

        para_text = (state.paragraphs[idx].text or "").lower()
        if search and search[:40] in para_text:
            return idx

    for idx in range(unit.start_paragraph_index + 1, unit.end_paragraph_index + 1):
        if idx < 0 or idx >= len(state.paragraphs):
            continue

        para = state.paragraphs[idx]
        if not para.is_heading and len((para.text or "").strip()) >= 60:
            return idx

    return min(unit.start_paragraph_index + 1, unit.end_paragraph_index)


# ───────────────────────────────────────────────
# Worker Function (Parallel)
# ───────────────────────────────────────────────

def _process_unit_accuracy(client, use_ollama, state, unit):
    try:
        prompt = _build_accuracy_prompt(unit.title, unit.body_text[:10000])

        if use_ollama:
            raw_findings = call_ollama_json(
                OLLAMA_ACCURACY_MODEL,
                "You produce strict JSON only.",
                prompt
            ).get("findings", [])
        else:
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

        return unit_findings[:MAX_ACCURACY_FINDINGS_PER_UNIT]

    except Exception:
        return []


# ───────────────────────────────────────────────
# Main Function
# ───────────────────────────────────────────────

def run_accuracy_check(state):
    if not ENABLE_ACCURACY_CHECK:
        return state

    mistral_key = os.environ.get("MISTRAL_API_KEY")
    use_ollama  = not mistral_key and ENABLE_OLLAMA_FALLBACK

    if not mistral_key and not use_ollama:
        return state

    if not getattr(state, "units", []):
        return state

    client = Mistral(api_key=mistral_key) if mistral_key else None

    units_to_check = [
        u for u in state.units
        if u.body_text and len(u.body_text.strip()) >= ACCURACY_MIN_UNIT_TEXT_CHARS
    ][:MAX_UNITS_FOR_ACCURACY]

    all_findings = []

    MAX_WORKERS = 4

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [
            executor.submit(_process_unit_accuracy, client, use_ollama, state, unit)
            for unit in units_to_check
        ]

        for future in as_completed(futures):
            all_findings.extend(future.result())

    state.accuracy_findings = all_findings
    return state
