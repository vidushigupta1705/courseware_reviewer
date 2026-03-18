from pathlib import Path
from statistics import mean

from config import (
    MIN_TEXT_CHARS_FOR_USEFUL_OCR,
    MIN_AVG_OCR_CONFIDENCE,
    MIN_CODE_HINT_SCORE,
)
from models import OCRLine, OCRResult
from ocr.ocr_engine import run_ocr
from ocr.image_classifier import classify_by_ocr_preview, fallback_classify_from_filename, score_code_like_text
from ocr.text_cleaner import clean_ocr_text, merge_lines_to_text
from ocr.code_cleaner import clean_code_ocr_text, looks_like_code


def process_single_image_ocr(image_info):
    raw_lines = run_ocr(Path(image_info.output_path))

    if not raw_lines:
        return OCRResult(
            image_filename=image_info.filename,
            image_path=image_info.output_path,
            image_type="unknown",
            lines=[],
            merged_text="",
            avg_confidence=0.0,
            status="no_text_detected",
        )

    texts = [t for t, _ in raw_lines]
    confidences = [c for _, c in raw_lines]
    preview_text = "\n".join(texts)
    image_type = classify_by_ocr_preview(preview_text)

    if image_type == "unknown":
        image_type = fallback_classify_from_filename(image_info.filename)

    merged = merge_lines_to_text(texts)

    code_hint_score = score_code_like_text(merged)
    if image_type != "code" and (looks_like_code(merged) or code_hint_score >= MIN_CODE_HINT_SCORE):
        image_type = "code"

    if image_type == "code":
        merged = clean_code_ocr_text(merged)
    else:
        merged = clean_ocr_text(merged)

    avg_conf = round(mean(confidences), 4) if confidences else 0.0

    if len(merged) < MIN_TEXT_CHARS_FOR_USEFUL_OCR:
        status = "low_text"
    elif avg_conf < MIN_AVG_OCR_CONFIDENCE:
        status = "low_confidence"
    else:
        status = "success"

    return OCRResult(
        image_filename=image_info.filename,
        image_path=image_info.output_path,
        image_type=image_type,
        lines=[OCRLine(text=t, confidence=float(c)) for t, c in raw_lines],
        merged_text=merged,
        avg_confidence=avg_conf,
        status=status,
    )
