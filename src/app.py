"""
app.py  —  Streamlit UI for the Courseware Review System
Run with:  streamlit run app.py
"""
import logging

import os
import sys
import shutil
import tempfile
import traceback
from pathlib import Path

import streamlit as st

# ── logging setup ─────────────────────────────────────────
LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# ── make src/ importable ──────────────────────────────────────────────────────
SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

# ── page config  (must be the very first Streamlit call) ─────────────────────
st.set_page_config(
    page_title="Courseware Reviewer",
    page_icon="📄",
    layout="wide",
)

# ── load .env if present (for API keys) ──────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv is optional; keys can be set as real env vars


# ─────────────────────────────────────────────────────────────────────────────
# Startup dependency check
# ─────────────────────────────────────────────────────────────────────────────

def _check_dependencies():
    """
    Returns (all_required_ok: bool, results: list[dict]).
    Each dict has keys: label, status ("ok" | "warn" | "error"), detail.
    """
    results = []

    def add(label, status, detail):
        results.append({"label": label, "status": status, "detail": detail})

    # Required: MISTRAL_API_KEY (or Ollama fallback)
    import requests as _req
    from config import ENABLE_OLLAMA_FALLBACK, OLLAMA_BASE_URL

    mistral_key = os.environ.get("MISTRAL_API_KEY", "").strip()
    if mistral_key:
        add("Mistral API Key", "ok", "Found in environment")
    else:
        add("Mistral API Key", "warn",
            "Not set — checking Ollama fallback.")

    if ENABLE_OLLAMA_FALLBACK:
        try:
            r = _req.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=3)
            if r.ok:
                models = [m["name"] for m in r.json().get("models", [])]
                add("Ollama (fallback LLM)", "ok",
                    f"Running. Models: {', '.join(models) or 'none pulled — run: ollama pull llama3'}")
            else:
                add("Ollama (fallback LLM)", "warn", "Reachable but returned an error.")
        except Exception:
            if not mistral_key:
                add("Ollama (fallback LLM)", "error",
                    "Not reachable and MISTRAL_API_KEY is missing — LLM features will not run. "
                    "Install Ollama from https://ollama.com and run: ollama pull llama3")
            else:
                add("Ollama (fallback LLM)", "warn",
                    "Not reachable — Mistral will be used instead.")

    # Optional: WINSTON_API_KEY
    if os.environ.get("WINSTON_API_KEY", "").strip():
        add("Winston AI Key (plagiarism)", "ok", "Found in environment")
    else:
        add("Winston AI Key (plagiarism)", "warn",
            "Not set — plagiarism scanning will be skipped.")

    # Optional: SERPAPI_KEY
    if os.environ.get("SERPAPI_KEY", "").strip():
        add("SerpAPI Key (image lookup)", "ok", "Found in environment")
    else:
        add("SerpAPI Key (image lookup)", "warn",
            "Not set — image source lookup will be skipped.")

    # External binary: tesseract
    if shutil.which("tesseract"):
        add("Tesseract OCR", "ok", f"Found at {shutil.which('tesseract')}")
    else:
        add("Tesseract OCR", "warn",
            "Not found — OCR on embedded images will be skipped.")

    # External binary: graphviz dot
    if shutil.which("dot"):
        add("Graphviz (dot)", "ok", f"Found at {shutil.which('dot')}")
    else:
        add("Graphviz (dot)", "warn",
            "Not found — diagram generation will be skipped.")

    # Python packages
    pkg_checks = [
        ("docx",       "python-docx",  True),
        ("mistralai",  "mistralai",    True),
        ("PIL",        "Pillow",       True),
        ("pytesseract","pytesseract",  False),
        ("graphviz",   "graphviz",     False),
        ("rapidfuzz",  "rapidfuzz",    True),
    ]
    for import_name, pkg_label, required in pkg_checks:
        try:
            __import__(import_name)
            add(f"Package: {pkg_label}", "ok", "Installed")
        except ImportError:
            level = "error" if required else "warn"
            add(f"Package: {pkg_label}", level,
                f"Not installed — run: pip install {pkg_label}")

    all_required_ok = all(r["status"] != "error" for r in results)
    return all_required_ok, results


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline runner (wraps run_step11.process_file)
# ─────────────────────────────────────────────────────────────────────────────

def _run_pipeline(uploaded_docx_path: Path, progress_bar, status_text):
    """
    Imports and runs the pipeline from src/run_step11.py.
    Returns (review_path, fixed_path, state) or raises on fatal error.
    """
    from config import (
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
    from checkpoint_manager import save_checkpoint

    stages = [
        ("Loading document",            lambda s: load_docx_state(uploaded_docx_path, EXTRACTED_IMAGES_DIR), "loaded"),
        ("Running OCR",                 lambda s: run_ocr_for_document(s),             "ocr_done"),
        ("Analysing image metadata",    lambda s: analyze_image_neighbors(s),          "image_metadata_done"),
        ("Checking duplicates",         lambda s: analyze_duplicates(s),               "duplicates_done"),
        ("Winston similarity scan",     lambda s: analyze_winston_similarity(s),       "winston_done"),
        ("Building units",              lambda s: build_units(s),                      "units_done"),
        ("Recommending diagrams",       lambda s: analyze_diagram_recommendations(s),  "diagram_recommend_done"),
        ("Building visual specs",       lambda s: build_visual_specs(s),               "visual_spec_done"),
        ("Rendering advanced visuals",  lambda s: render_advanced_visuals(s),          "advanced_visual_done"),
        ("Analysing tables & code",     lambda s: analyze_tables_and_code(s),          "table_code_done"),
        ("Checking content accuracy",   lambda s: run_accuracy_check(s),               "accuracy_done"),
        ("Running deterministic checks",lambda s: run_all_checks_step11(s),            "checks_done"),
        ("LLM rewriting", lambda s: run_llm_rewrite(s) if getattr(s, "issues", None) else s, "rewrite_done"),
        ("Generating quizzes", lambda s: generate_quizzes_for_units(s) if getattr(s, "units", None) else s, "quiz_done"),
    ]

    state = None
    total = len(stages)
    skipped_stages = []

    # Stage 0: load is mandatory — fail hard if this breaks
    label, fn, checkpoint_name = stages[0]
    status_text.text(f"⏳ {label}…")
    state = fn(None)  # load doesn't use state arg
    save_checkpoint(state, checkpoint_name)
    logging.info(f"[PIPELINE] Completed stage: {label}")
    progress_bar.progress(1 / total)

    # ── Lightweight backup/restore helpers ──────────────────────────────────
    # Only mutable fields are backed up — paragraphs, images, and other
    # ingest-time data never change so they don't need to be copied.
    # This replaces deepcopy(state) which copied 50-100MB per stage.
    _MUTABLE_FIELDS = [
        "issues", "duplicate_findings", "retrieval_findings",
        "rewrite_suggestions", "units", "ocr_results",
        "accuracy_findings", "diagram_recommendations",
        "generated_diagrams", "generated_visuals", "visual_specs",
        "table_findings", "code_findings",
        "duplicate_findings_truncated", "duplicate_findings_total",
        "image_sources_found",
    ]

    def _backup(s):
        return {
            f: list(getattr(s, f, None) or [])
            for f in _MUTABLE_FIELDS
            if hasattr(s, f)
        }

    def _restore(s, bk):
        for f, v in bk.items():
            setattr(s, f, v)
        return s

    # Remaining stages: safe execution with rollback
    for i, (label, fn, checkpoint_name) in enumerate(stages[1:], start=2):
        status_text.text(f"⏳ {label}…")

        backup = _backup(state)  # lightweight field-level backup

        try:
            result = fn(state)

            # Support both return-style and in-place mutation
            if result is not None:
                state = result

            save_checkpoint(state, checkpoint_name)
            logging.info(f"[PIPELINE] Completed stage: {label}")

        except Exception as exc:
            logging.exception(f"[PIPELINE ERROR] Stage failed: {label}")

            skipped_stages.append((label, str(exc)))

            # Restore only the mutable fields — ingest data is untouched
            _restore(state, backup)

        progress_bar.progress(i / total)

    # Build output documents
    status_text.text("⏳ Building output documents…")
    base = uploaded_docx_path.stem
    review_path = REVIEW_COMMENTS_DIR / f"{base}_Review Comments.docx"
    fixed_path  = FINAL_FIXED_DIR     / f"{base}_Final Fixed.docx"

    REVIEW_COMMENTS_DIR.mkdir(parents=True, exist_ok=True)
    FINAL_FIXED_DIR.mkdir(parents=True, exist_ok=True)

    build_review_comments_doc(uploaded_docx_path, state, review_path)
    build_final_fixed_doc(uploaded_docx_path, state, fixed_path)

    progress_bar.progress(1.0)
    status_text.text("✅ Done!")

    return review_path, fixed_path, state, skipped_stages


# ─────────────────────────────────────────────────────────────────────────────
# Issue table helpers
# ─────────────────────────────────────────────────────────────────────────────

SEVERITY_EMOJI = {
    "high":   "🔴",
    "medium": "🟠",
    "low":    "🟡",
    "info":   "🔵",
}

ISSUE_TYPE_TAG = {
    "grammar":                      "[Grammar]",
    "spacing":                      "[Format]",
    "formatting_body":              "[Format]",
    "formatting_heading":           "[Format]",
    "heading_hierarchy":            "[Structure]",
    "navigation_pane":              "[Structure]",
    "duplicate_paragraph":          "[Duplicate]",
    "near_duplicate_paragraph":     "[Duplicate]",
    "possible_duplicate_topic":     "[Duplicate]",
    "repeated_heading":             "[Structure]",
    "image_caption_missing":        "[Image]",
    "image_source_missing":         "[Image]",
    "figure_number_missing":        "[Image]",
    "ocr_low_confidence":           "[OCR]",
    "content_accuracy":             "[Accuracy]",
    "quiz_missing":                 "[Quiz]",
    "diagram_recommended":          "[Diagram]",
    "diagram_generated":            "[Diagram]",
    "diagram_generation_failed":    "[Diagram]",
    "advanced_visual_generated":    "[Visual]",
    "possible_ibm_similarity":      "[Plagiarism]",
    "possible_open_source_similarity": "[Plagiarism]",
    "winston_scan_error":           "[Plagiarism]",
}


def _tag(issue_type: str) -> str:
    return ISSUE_TYPE_TAG.get(issue_type, f"[{issue_type.replace('_', ' ').title()}]")


def _severity_order(sev: str) -> int:
    return {"high": 0, "medium": 1, "low": 2, "info": 3}.get(sev, 4)


# ─────────────────────────────────────────────────────────────────────────────
# UI — Sidebar: dependency status
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("📄 Courseware Reviewer")
    st.markdown("---")
    st.subheader("🔧 System Status")

    deps_ok, dep_results = _check_dependencies()

    for dep in dep_results:
        icon = {"ok": "✅", "warn": "⚠️", "error": "❌"}[dep["status"]]
        with st.expander(f"{icon} {dep['label']}"):
            st.caption(dep["detail"])

    if not deps_ok:
        st.error("One or more required items are missing. See details above.")
    else:
        st.success("All required checks passed.")

    st.markdown("---")
    st.caption("Upload a .docx file on the right to begin.")


# ─────────────────────────────────────────────────────────────────────────────
# UI — Main area
# ─────────────────────────────────────────────────────────────────────────────

st.title("📄 Courseware Review System")
st.markdown(
    "Upload a Word document (.docx) to run the full review pipeline. "
    "You will get two output files: a **Review Comments** doc and a **Final Fixed** doc."
)


uploaded_file = st.file_uploader("Choose a .docx file", type=["docx"])

# ── Session state init ────────────────────────────────────────────────────────
# All pipeline results are stored here so Streamlit reruns (triggered by
# download button clicks, filter changes etc.) never re-execute the pipeline.
if "pipeline_done" not in st.session_state:
    st.session_state.pipeline_done = False
if "review_path" not in st.session_state:
    st.session_state.review_path = None
if "fixed_path" not in st.session_state:
    st.session_state.fixed_path = None
if "state" not in st.session_state:
    st.session_state.state = None
if "skipped" not in st.session_state:
    st.session_state.skipped = []
if "last_filename" not in st.session_state:
    st.session_state.last_filename = None

# Reset if a different file is uploaded
if uploaded_file is not None and uploaded_file.name != st.session_state.last_filename:
    st.session_state.pipeline_done = False
    st.session_state.review_path = None
    st.session_state.fixed_path = None
    st.session_state.state = None
    st.session_state.skipped = []
    st.session_state.last_filename = uploaded_file.name

if uploaded_file is not None:
    if not deps_ok:
        st.error(
            "Cannot run the pipeline — one or more required dependencies are missing. "
            "Check the sidebar for details."
        )
        st.stop()

    st.divider()
    st.subheader(f"📂 File: {uploaded_file.name}")

    # Only show Run button if pipeline hasn't completed for this file yet
    if not st.session_state.pipeline_done:
        run_button = st.button("▶️  Run Review Pipeline", type="primary")

        if run_button:
            from config import INPUT_DIR
            INPUT_DIR.mkdir(parents=True, exist_ok=True)

            input_path = INPUT_DIR / uploaded_file.name
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            st.divider()
            st.subheader("⏳ Pipeline Progress")
            progress_bar = st.progress(0)
            status_text  = st.empty()

            try:
                review_path, fixed_path, state, skipped = _run_pipeline(
                    input_path, progress_bar, status_text
                )
                # Store everything in session state — survives all future reruns
                st.session_state.pipeline_done = True
                st.session_state.review_path   = review_path
                st.session_state.fixed_path    = fixed_path
                st.session_state.state         = state
                st.session_state.skipped       = skipped
            except Exception:
                st.error(f"Pipeline failed during document loading:\n\n```\n{traceback.format_exc()}\n```")
                st.stop()

    # ── Results section — shown whenever pipeline_done is True ───────────────
    if st.session_state.pipeline_done:
        state       = st.session_state.state
        review_path = st.session_state.review_path
        fixed_path  = st.session_state.fixed_path
        skipped     = st.session_state.skipped

        # Reset button
        if st.button("🔄  Run on a different file"):
            st.session_state.pipeline_done = False
            st.session_state.last_filename = None
            st.rerun()

        # ── Skipped stages warning ────────────────────────────────────────────
        if skipped:
            with st.expander(f"⚠️  {len(skipped)} stage(s) skipped due to errors — click to see details"):
                for stage_name, err_msg in skipped:
                    st.markdown(f"**{stage_name}**")
                    st.code(err_msg)

        st.divider()

        # ── Summary metrics ───────────────────────────────────────────────────
        st.subheader("📊 Review Summary")

        issues   = getattr(state, "issues", [])
        units    = getattr(state, "units", [])
        rewrites = getattr(state, "rewrite_suggestions", [])
        visuals  = getattr(state, "generated_visuals", [])
        accuracy = getattr(state, "accuracy_findings", [])

        high_count   = sum(1 for i in issues if i.severity == "high")
        medium_count = sum(1 for i in issues if i.severity == "medium")
        low_count    = sum(1 for i in issues if i.severity in ("low", "info"))

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total Issues",   len(issues))
        col2.metric("🔴 High",        high_count)
        col3.metric("🟠 Medium",      medium_count)
        col4.metric("🟡 Low / Info",  low_count)
        col5.metric("Units Detected", len(units))

        col6, col7, col8 = st.columns(3)
        col6.metric("Rewrite Suggestions", len(rewrites))
        col7.metric("Visuals Generated",   len(visuals))
        col8.metric("Accuracy Findings",   len(accuracy))

        st.divider()

        # ── Download buttons ──────────────────────────────────────────────────
        st.subheader("📥 Download Outputs")
        dl_col1, dl_col2 = st.columns(2)

        if review_path and Path(review_path).exists():
            with open(review_path, "rb") as f:
                dl_col1.download_button(
                    label="⬇️  Review Comments (.docx)",
                    data=f.read(),
                    file_name=Path(review_path).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        else:
            dl_col1.warning("Review Comments file was not created.")

        if fixed_path and Path(fixed_path).exists():
            with open(fixed_path, "rb") as f:
                dl_col2.download_button(
                    label="⬇️  Final Fixed (.docx)",
                    data=f.read(),
                    file_name=Path(fixed_path).name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        else:
            dl_col2.warning("Final Fixed file was not created.")

        st.divider()

        # ── Issues table ──────────────────────────────────────────────────────
        if issues:
            st.subheader("🔍 All Issues Found")

            severity_filter = st.multiselect(
                "Filter by severity",
                options=["high", "medium", "low", "info"],
                default=["high", "medium", "low", "info"],
            )

            all_types = sorted(set(i.issue_type for i in issues))
            type_filter = st.multiselect(
                "Filter by type",
                options=all_types,
                default=all_types,
            )

            filtered = [
                i for i in issues
                if i.severity in severity_filter and i.issue_type in type_filter
            ]
            filtered.sort(key=lambda i: _severity_order(i.severity))

            st.caption(f"Showing {len(filtered)} of {len(issues)} issues")

            for issue in filtered:
                sev_icon = SEVERITY_EMOJI.get(issue.severity, "⚪")
                tag      = _tag(issue.issue_type)
                header   = f"{sev_icon} {tag}  —  {issue.location}"

                with st.expander(header):
                    st.markdown(f"**Severity:** {issue.severity.upper()}")
                    st.markdown(f"**Issue:** {issue.message}")
                    if issue.suggested_fix:
                        st.markdown(f"**Suggested Fix:** {issue.suggested_fix}")
                    if issue.paragraph_index is not None:
                        st.caption(f"Paragraph index: {issue.paragraph_index}")
        else:
            st.success("No issues found — the document looks clean!")

        # ── Units detected ────────────────────────────────────────────────────
        if units:
            st.divider()
            st.subheader("📚 Units / Chapters Detected")
            for unit in units:
                has_quiz = "✅ Quiz present" if unit.quiz_present else (
                    f"📝 {len(unit.generated_quiz)} question(s) generated"
                    if unit.generated_quiz else "❌ No quiz"
                )
                st.markdown(f"- **{unit.title}** — {has_quiz}")
