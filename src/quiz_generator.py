import json
import os

from mistralai.client import Mistral

from config import ENABLE_QUIZ_GENERATION, QUIZ_MODEL, MAX_UNITS_FOR_QUIZ
from models import QuizItem


def _build_quiz_prompt(unit_title: str, unit_text: str) -> str:
    return f"""
You are generating a courseware quiz.

Generate exactly 15 questions for the following chapter/unit.

Rules:
1. Use only the provided chapter/unit content.
2. Do not add facts not present in the text.
3. Return strict JSON only.
4. JSON must contain one key: quiz
5. quiz must be a list of exactly 15 objects in this exact order:
   - 3 MCQ
   - 3 Fill in the Blanks
   - 3 True/False
   - 3 Two-mark
   - 3 Four-mark
6. Each object must contain:
   - qtype
   - question
   - options (empty list unless MCQ)
   - answer
7. For MCQ, provide exactly 4 options.

Chapter/Unit title:
{unit_title}

Chapter/Unit content:
{unit_text}
""".strip()


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


def _validate_quiz_items(items):
    """
    Strict validation:
    must be exactly 15 in required order
    """
    if len(items) != 15:
        return False

    expected_order = (
        ["MCQ"] * 3
        + ["Fill in the Blanks"] * 3
        + ["True/False"] * 3
        + ["Two-mark"] * 3
        + ["Four-mark"] * 3
    )

    normalized = []
    for item in items:
        qtype = (item.get("qtype", "") or "").strip()
        normalized.append(qtype)

    return normalized == expected_order


def generate_quizzes_for_units(state):
    if not ENABLE_QUIZ_GENERATION:
        return state

    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        return state

    client = Mistral(api_key=api_key)
    final_units = []

    for unit in state.units[:MAX_UNITS_FOR_QUIZ]:
        if unit.quiz_present:
            final_units.append(unit)
            continue

        if unit.end_paragraph_index <= unit.heading_paragraph_index:
            final_units.append(unit)
            continue

        if not unit.body_text or len(unit.body_text.strip()) < 250:
            final_units.append(unit)
            continue

        try:
            prompt = _build_quiz_prompt(unit.title, unit.body_text[:12000])
            result = _call_mistral_quiz(client, prompt)
            raw_items = result.get("quiz", [])

            if not _validate_quiz_items(raw_items):
                final_units.append(unit)
                continue

            quiz_items = []
            for item in raw_items:
                quiz_items.append(
                    QuizItem(
                        qtype=(item.get("qtype", "") or "").strip(),
                        question=(item.get("question", "") or "").strip(),
                        options=item.get("options", []) or [],
                        answer=(item.get("answer", "") or "").strip(),
                    )
                )

            unit.generated_quiz = quiz_items
            final_units.append(unit)

        except Exception:
            final_units.append(unit)

    state.units = final_units
    return state