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
from image_metadata_analysis import analyze_image_neighbors
from duplicate_analysis import analyze_duplicates
from winston_similarity import analyze_winston_similarity
from unit_builder import build_units
from deterministic_checks import run_all_checks_step7
from llm_rewrite import run_llm_rewrite
from accuracy_checker import run_accuracy_check
from quiz_generator import generate_quizzes_for_units
from docx_writer import build_review_comments_doc, build_final_fixed_doc
from utils import save_json
from checkpoint_manager import save_checkpoint, load_checkpoint


STAGES = [
    "loaded",
    "ocr_done",
    "image_metadata_done",
    "duplicates_done",
    "winston_done",
    "units_done",
    "accuracy_done",
    "checks_done",
    "rewrite_done",
    "quiz_done",
]


def _find_latest_checkpoint(base_filename: str):
    for stage in reversed(STAGES):
        state = load_checkpoint(base_filename, stage)
        if state is not None:
            return stage, state
    return None, None


def process_file(input_file: Path, resume=True):
    base_filename = input_file.stem

    current_stage, state = (None, None)
    if resume:
        current_stage, state = _find_latest_checkpoint(base_filename)

    if state is None:
        state = load_docx_state(input_file, EXTRACTED_IMAGES_DIR)
        save_checkpoint(state, "loaded")
        current_stage = "loaded"

    if current_stage == "loaded":
        state = run_ocr_for_document(state)
        save_checkpoint(state, "ocr_done")
        current_stage = "ocr_done"

    if current_stage == "ocr_done":
        state = analyze_image_neighbors(state)
        save_checkpoint(state, "image_metadata_done")
        current_stage = "image_metadata_done"

    if current_stage == "image_metadata_done":
        state = analyze_duplicates(state)
        save_checkpoint(state, "duplicates_done")
        current_stage = "duplicates_done"

    if current_stage == "duplicates_done":
        state = analyze_winston_similarity(state)
        save_checkpoint(state, "winston_done")
        current_stage = "winston_done"

    if current_stage == "winston_done":
        state = build_units(state)
        save_checkpoint(state, "units_done")
        current_stage = "units_done"

    if current_stage == "units_done":
        state = run_accuracy_check(state)
        save_checkpoint(state, "accuracy_done")
        current_stage = "accuracy_done"

    if current_stage == "accuracy_done":
        state = run_all_checks_step7(state)
        save_checkpoint(state, "checks_done")
        current_stage = "checks_done"

    if current_stage == "checks_done":
        state = run_llm_rewrite(state)
        save_checkpoint(state, "rewrite_done")
        current_stage = "rewrite_done"

    if current_stage == "rewrite_done":
        state = generate_quizzes_for_units(state)
        save_checkpoint(state, "quiz_done")
        current_stage = "quiz_done"

    review_comments_path = REVIEW_COMMENTS_DIR / f"{state.base_filename}_Review Comments.docx"
    final_fixed_path = FINAL_FIXED_DIR / f"{state.base_filename}_Final Fixed.docx"
    json_path = INTERMEDIATE_JSON_DIR / f"{state.base_filename}_step7_checkpointed_state.json"

    build_review_comments_doc(input_file, state, review_comments_path)
    build_final_fixed_doc(input_file, state, final_fixed_path)
    save_json(state.to_dict(), json_path)

    print("Checkpointed Step 7 completed successfully.")
    print(f"Review comments file: {review_comments_path}")
    print(f"Final fixed file: {final_fixed_path}")
    print(f"State JSON: {json_path}")
    print(f"Units detected: {len(getattr(state, 'units', []))}")
    print(f"Rewrite suggestions: {len(getattr(state, 'rewrite_suggestions', []))}")
    print(f"Winston findings: {len(getattr(state, 'retrieval_findings', []))}")


if __name__ == "__main__":
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None

    if input_arg:
        file_path = Path(input_arg)
    else:
        docx_files = list(INPUT_DIR.glob("*.docx"))
        if not docx_files:
            raise FileNotFoundError(f"No DOCX files found in {INPUT_DIR}")
        file_path = docx_files[0]

    process_file(file_path, resume=True)