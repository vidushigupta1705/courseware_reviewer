import sys
from pathlib import Path

from config import (
    INPUT_DIR,
    REVIEW_COMMENTS_DIR,
    FINAL_FIXED_DIR,
    INTERMEDIATE_JSON_DIR,
    EXTRACTED_IMAGES_DIR,
)
from docx_ingest import load_docx_state
from ocr_pipeline import run_ocr_for_document
from deterministic_checks import run_all_checks_with_ocr
from docx_writer import build_review_comments_doc, build_final_fixed_doc
from utils import save_json


def process_file(input_file: Path):
    state = load_docx_state(input_file, EXTRACTED_IMAGES_DIR)
    state = run_ocr_for_document(state)
    state = run_all_checks_with_ocr(state)

    review_comments_path = REVIEW_COMMENTS_DIR / f"{state.base_filename}_Review Comments.docx"
    final_fixed_path = FINAL_FIXED_DIR / f"{state.base_filename}_Final Fixed.docx"
    json_path = INTERMEDIATE_JSON_DIR / f"{state.base_filename}_step2_state.json"

    build_review_comments_doc(input_file, state, review_comments_path)
    build_final_fixed_doc(input_file, state, final_fixed_path)
    save_json(state.to_dict(), json_path)

    print("Step 2 completed successfully.")
    print(f"Review comments file: {review_comments_path}")
    print(f"Final fixed file: {final_fixed_path}")
    print(f"State JSON: {json_path}")
    print(f"Extracted images count: {len(state.images)}")
    print(f"OCR results count: {len(state.ocr_results)}")
    print(f"Issues detected: {len(state.issues)}")


if __name__ == "__main__":
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None

    if input_arg:
        file_path = Path(input_arg)
    else:
        docx_files = list(INPUT_DIR.glob("*.docx"))
        if not docx_files:
            raise FileNotFoundError(f"No DOCX files found in {INPUT_DIR}")
        file_path = docx_files[0]

    process_file(file_path)
