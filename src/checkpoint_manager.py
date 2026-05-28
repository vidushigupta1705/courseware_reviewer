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

# Single source of truth — must stay in sync between save and load.
_DYNAMIC_ATTRS = (
    "duplicate_findings",
    "retrieval_findings",
    "table_findings",
    "code_findings",
    "accuracy_findings",        # was missing from save_checkpoint
    "diagram_recommendations",
    "generated_diagrams",
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

    for attr in _DYNAMIC_ATTRS:
        if hasattr(state, attr):
            payload[attr] = _to_plain(getattr(state, attr))

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

def _load_visual_specs(items):
    from models import VisualSpec, VisualNode, VisualEdge
    specs = []
    for item in items:
        nodes = [VisualNode(**n) for n in item.get("nodes", [])]
        edges = [VisualEdge(**e) for e in item.get("edges", [])]
        item_copy = dict(item)
        item_copy["nodes"] = nodes
        item_copy["edges"] = edges
        specs.append(VisualSpec(**item_copy))
    return specs

def _load_generated_visuals(items):
    from models import GeneratedVisual
    return [GeneratedVisual(**x) for x in items]


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
        visual_specs=_load_visual_specs(data.get("visual_specs", [])),
        generated_visuals=_load_generated_visuals(data.get("generated_visuals", [])),
    )

    for attr in _DYNAMIC_ATTRS:
        setattr(state, attr, data.get(attr, []))   # safe default — no KeyError

    return state
