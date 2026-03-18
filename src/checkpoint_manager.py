import json
from pathlib import Path
from dataclasses import asdict, is_dataclass

from config import CHECKPOINT_DIR, ENABLE_CHECKPOINTS
from models import (
    DocumentState,
    ParagraphInfo,
    ImageInfo,
    OCRLine,
    OCRResult,
    RewriteSuggestion,
    QuizItem,
    UnitInfo,
    ReviewIssue,
)


def _to_plain(obj):
    if is_dataclass(obj):
        return asdict(obj)
    if isinstance(obj, list):
        return [_to_plain(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_plain(v) for k, v in obj.items()}
    return obj


def save_checkpoint(state: DocumentState, stage_name: str):
    if not ENABLE_CHECKPOINTS:
        return

    CHECKPOINT_DIR.mkdir(parents=True, exist_ok=True)
    path = CHECKPOINT_DIR / f"{state.base_filename}__{stage_name}.json"

    payload = state.to_dict()

    # include optional dynamic attrs if present
    if hasattr(state, "duplicate_findings"):
        payload["duplicate_findings"] = getattr(state, "duplicate_findings")
    if hasattr(state, "retrieval_findings"):
        payload["retrieval_findings"] = getattr(state, "retrieval_findings")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)


def _load_paragraphs(items):
    return [ParagraphInfo(**x) for x in items]


def _load_images(items):
    return [ImageInfo(**x) for x in items]


def _load_ocr_results(items):
    results = []
    for item in items:
        lines = [OCRLine(**ln) for ln in item.get("lines", [])]
        item_copy = dict(item)
        item_copy["lines"] = lines
        results.append(OCRResult(**item_copy))
    return results


def _load_rewrite_suggestions(items):
    return [RewriteSuggestion(**x) for x in items]


def _load_units(items):
    units = []
    for item in items:
        quiz_items = [QuizItem(**q) for q in item.get("generated_quiz", [])]
        item_copy = dict(item)
        item_copy["generated_quiz"] = quiz_items
        units.append(UnitInfo(**item_copy))
    return units


def _load_issues(items):
    return [ReviewIssue(**x) for x in items]


def load_checkpoint(base_filename: str, stage_name: str):
    path = CHECKPOINT_DIR / f"{base_filename}__{stage_name}.json"
    if not path.exists():
        return None

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    state = DocumentState(
        input_file=data["input_file"],
        base_filename=data["base_filename"],
        paragraphs=_load_paragraphs(data.get("paragraphs", [])),
        images=_load_images(data.get("images", [])),
        ocr_results=_load_ocr_results(data.get("ocr_results", [])),
        rewrite_suggestions=_load_rewrite_suggestions(data.get("rewrite_suggestions", [])),
        units=_load_units(data.get("units", [])),
        issues=_load_issues(data.get("issues", [])),
    )

    if "duplicate_findings" in data:
        state.duplicate_findings = data["duplicate_findings"]
    if "retrieval_findings" in data:
        state.retrieval_findings = data["retrieval_findings"]

    return state