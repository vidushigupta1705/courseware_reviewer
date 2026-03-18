import json
import re
from pathlib import Path


def pt_to_float(size_obj):
    if size_obj is None:
        return None
    try:
        return round(size_obj.pt, 2)
    except Exception:
        return None


def normalize_text(text: str) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def has_extra_spacing(text: str) -> bool:
    if not text:
        return False

    patterns = [
        r"  +",
        r"\s+[,.;:!?]",
        r"[\t\r\f\v]",
    ]
    return any(re.search(p, text) for p in patterns)


def clean_spacing(text: str) -> str:
    if not text:
        return text
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def save_json(data, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
