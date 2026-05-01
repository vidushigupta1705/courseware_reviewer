import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

from mistralai.client import Mistral

from config import ENABLE_QUIZ_GENERATION, QUIZ_MODEL, MAX_UNITS_FOR_QUIZ
from models import QuizItem


# ───────────────────────────────────────────────
# Prompt Builder
# ───────────────────────────────────────────────

def _build_quiz_prompt(unit_title: str, unit_text: str) -> str:
    return f"""
You are generating a courseware quiz.

Generate exactly 15 questions for the following chapter/unit.

Rules:
1. Use only the provided content.
2. Do not add external facts.
3. Return strict JSON only.
4. JSON must contain key: quiz
5. quiz must contain exactly 15 questions:
   - 3 MCQ
   - 3 Fill in the Blanks
   - 3 True/False
   - 3 Two-mark
   - 3 Four-mark
6. Each object must include:
   - qtype
   - question
   - options (empty unless MCQ)
   - answer
7. MCQs must have exactly 4 options.

Title:
{unit_title}

Content:
{unit_text}
""".strip()


# ───────────────────────────────────────────────
# LLM Call
# ───────────────────────────────────────────────

def _call_mistral_quiz(client: Mistral, prompt: str):
    response = client.chat.complete(
        model=QUIZ_MODEL,
        messages=[
            {"role": "system", "content": "You produce strict JSON only."},
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.1,
        random_seed=11,
    )

    content = response.choices[0].message.content

    if isinstance(content, list):
        content = "".join(
            chunk.get("text", "") if isinstance(chunk, dict) else str(chunk)
            for chunk in content
        )

    return json.loads(content)


# ───────────────────────────────────────────────
# Validation
# ───────────────────────────────────────────────

def _validate_quiz_items(items):
    if len(items) != 15:
        return False

    expected_order = (
        ["MCQ"] * 3
        + ["Fill in the Blanks"] * 3
        + ["True/False"] * 3
        + ["Two-mark"] * 3
        + ["Four-mark"] * 3
    )

    normalized = [(item.get("qtype", "") or "").strip() for item in items]
    return normalized == expected_order


# ───────────────────────────────────────────────
# Worker Function (Parallel)
# ───────────────────────────────────────────────

def _process_unit(client, unit):
    if unit.quiz_present:
        return unit

    if unit.end_paragraph_index <= unit.heading_paragraph_index:
        return unit

    if not unit.body_text or len(unit.body_text.strip()) < 250:
        return unit

    try:
        prompt = _build_quiz_prompt(unit.title, unit.body_text[:12000])
        result = _call_mistral_quiz(client, prompt)
        raw_items = result.get("quiz", [])

        # Normalize common label variations before validating
        _LABEL_MAP = {
            "fill_in_the_blanks": "Fill in the Blanks",
            "fill in the blank":  "Fill in the Blanks",
            "true_false":         "True/False",
            "true or false":      "True/False",
            "two_mark":           "Two-mark",
            "two mark":           "Two-mark",
            "four_mark":          "Four-mark",
            "four mark":          "Four-mark",
            "mcq":                "MCQ",
            "multiple choice":    "MCQ",
        }
        for item in raw_items:
            raw = str(item.get("qtype") or "").strip().lower()
            if raw in _LABEL_MAP:
                item["qtype"] = _LABEL_MAP[raw]

        if not _validate_quiz_items(raw_items):
            print(f"[QUIZ VALIDATION FAILED] '{unit.title}' — qtypes returned: {[i.get('qtype') for i in raw_items]}")
            return unit

        quiz_items = [
            QuizItem(
                qtype=str(item.get("qtype") or "").strip(),
                question=str(item.get("question") or "").strip(),
                options=item.get("options", []) or [],
                answer=str(item.get("answer") or "").strip(),
            )
            for item in raw_items
        ]

        unit.generated_quiz = quiz_items
        return unit

    except Exception as e:
        print(f"[QUIZ ERROR] '{unit.title}': {type(e).__name__}: {e}")
        return unit


# ───────────────────────────────────────────────
# Main Function
# ───────────────────────────────────────────────

def generate_quizzes_for_units(state):
    if not ENABLE_QUIZ_GENERATION:
        return state

    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        return state

    if not getattr(state, "units", []):
        return state

    client = Mistral(api_key=api_key)

    units = state.units[:MAX_UNITS_FOR_QUIZ]
    results = []

    MAX_WORKERS = 4  # safe parallelism

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(_process_unit, client, unit) for unit in units]

        for future in as_completed(futures):
            results.append(future.result())

    # Maintain original order
    results.sort(key=lambda u: u.heading_paragraph_index)

    state.units = results
    return state
