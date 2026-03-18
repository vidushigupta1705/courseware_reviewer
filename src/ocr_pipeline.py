from pathlib import Path
import json

from config import OCR_RESULTS_DIR, OCR_DEBUG_TEXT_DIR
from ocr.ocr_router import process_single_image_ocr


def run_ocr_for_document(state):
    OCR_RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    OCR_DEBUG_TEXT_DIR.mkdir(parents=True, exist_ok=True)

    ocr_results = []

    for image_info in state.images:
        result = process_single_image_ocr(image_info)
        ocr_results.append(result)

        image_info.image_type = result.image_type
        image_info.ocr_text = result.merged_text
        image_info.ocr_confidence = result.avg_confidence
        image_info.ocr_status = result.status

        per_image_json = OCR_RESULTS_DIR / f"{state.base_filename}__{Path(image_info.filename).stem}.json"
        with open(per_image_json, "w", encoding="utf-8") as f:
            json.dump({
                "image_filename": result.image_filename,
                "image_path": result.image_path,
                "image_type": result.image_type,
                "avg_confidence": result.avg_confidence,
                "status": result.status,
                "lines": [{"text": l.text, "confidence": l.confidence} for l in result.lines],
                "merged_text": result.merged_text,
            }, f, indent=2, ensure_ascii=False)

        debug_txt = OCR_DEBUG_TEXT_DIR / f"{state.base_filename}__{Path(image_info.filename).stem}.txt"
        with open(debug_txt, "w", encoding="utf-8") as f:
            f.write(result.merged_text or "")

    state.ocr_results = ocr_results
    return state
