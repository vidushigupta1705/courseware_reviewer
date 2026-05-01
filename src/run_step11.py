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
from diagram_recommender import analyze_diagram_recommendations
from visual_spec_builder import build_visual_specs
from advanced_visual_renderer import render_advanced_visuals
from table_code_analysis import analyze_tables_and_code
from deterministic_checks import run_all_checks_step11
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
    "diagram_recommend_done",
    "visual_spec_done",
    "advanced_visual_done",
    "table_code_done",
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
    skipped = []  # collects (stage_label, error_message) for any skipped stages

    latest_completed, state = (None, None)
    if resume:
        latest_completed, state = _find_latest_checkpoint(base_filename)

    # Document loading is mandatory — if this fails, nothing else can run
    if state is None:
        state = load_docx_state(input_file, EXTRACTED_IMAGES_DIR)
        save_checkpoint(state, "loaded")
        latest_completed = "loaded"

    # Each tuple: (checkpoint_name_when_done, human_label, function)
    pipeline = [
        ("ocr_done", "OCR",
        lambda s: run_ocr_for_document(s) if getattr(s, "images", None) else s),
        ("image_metadata_done",   "Image metadata analysis",  analyze_image_neighbors),
        ("duplicates_done",       "Duplicate detection",      analyze_duplicates),
        ("winston_done",          "Winston similarity scan",  analyze_winston_similarity),
        ("units_done",            "Unit builder",             build_units),
        ("diagram_recommend_done", "Diagram recommendations",
        lambda s: analyze_diagram_recommendations(s) if getattr(s, "units", None) else s),
        ("visual_spec_done",      "Visual spec builder",      build_visual_specs),
        ("advanced_visual_done",  "Advanced visual renderer",
         lambda s: render_advanced_visuals(s) if getattr(s, "visual_specs", None) else s),
        ("table_code_done",       "Table & code analysis",    analyze_tables_and_code),
        ("accuracy_done",         "Accuracy check",           run_accuracy_check),
        ("checks_done",           "Deterministic checks",     run_all_checks_step11),
        ("rewrite_done", "LLM rewrite",
        lambda s: run_llm_rewrite(s) if getattr(s, "issues", None) else s),
        ("quiz_done", "Quiz generation",
        lambda s: generate_quizzes_for_units(s) if getattr(s, "units", None) else s),
    ]

    # Build a lookup so we can skip already-completed stages when resuming
    completed_stages = set(STAGES[: STAGES.index(latest_completed) + 1]) if latest_completed else set()

    for checkpoint_name, label, fn in pipeline:
        # Skip stages that were already completed in a previous run
        if checkpoint_name in completed_stages:
            print(f"   ✓ Skipping '{label}' — already completed (checkpoint found)")
            continue

        try:
            state = fn(state)
            save_checkpoint(state, checkpoint_name)
            completed_stages.add(checkpoint_name)
        except Exception as exc:
            skipped.append((label, str(exc)))
            # Mark as completed anyway so later stages still run
            completed_stages.add(checkpoint_name)

    # Report any skipped stages
    if skipped:
        print(f"\n⚠️  {len(skipped)} stage(s) skipped due to errors:")
        for stage_label, err_msg in skipped:
            print(f"   [{stage_label}] {err_msg}")
        print()

    review_comments_path = REVIEW_COMMENTS_DIR / f"{state.base_filename}_Review Comments.docx"
    final_fixed_path      = FINAL_FIXED_DIR     / f"{state.base_filename}_Final Fixed.docx"
    json_path             = INTERMEDIATE_JSON_DIR / f"{state.base_filename}_step11_state.json"

    build_review_comments_doc(input_file, state, review_comments_path)
    build_final_fixed_doc(input_file, state, final_fixed_path)
    save_json(state.to_dict(), json_path)

    print("Step 11 completed.")
    print(f"Review comments file : {review_comments_path}")
    print(f"Final fixed file     : {final_fixed_path}")
    print(f"State JSON           : {json_path}")
    print(f"Visual specs         : {len(getattr(state, 'visual_specs', []))}")
    print(f"Generated visuals    : {len(getattr(state, 'generated_visuals', []))}")
    print(f"Units detected       : {len(getattr(state, 'units', []))}")
    print(f"Rewrite suggestions  : {len(getattr(state, 'rewrite_suggestions', []))}")
    print(f"Stages skipped       : {len(skipped)}")

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
