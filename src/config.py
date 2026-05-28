from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# Directory layout
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR             = BASE_DIR / "input" / "raw_docx"
OUTPUT_DIR            = BASE_DIR / "output"
REVIEW_COMMENTS_DIR   = OUTPUT_DIR / "review_comments"
FINAL_FIXED_DIR       = OUTPUT_DIR / "final_fixed"
INTERMEDIATE_JSON_DIR = OUTPUT_DIR / "intermediate_json"
DEBUG_EXPORT_DIR      = OUTPUT_DIR / "debug_exports"
OCR_DEBUG_TEXT_DIR    = DEBUG_EXPORT_DIR / "ocr_text"

EXTRACTED_IMAGES_DIR  = BASE_DIR / "cache" / "extracted_images"
OCR_RESULTS_DIR       = BASE_DIR / "cache" / "ocr_results"
CHECKPOINT_DIR        = BASE_DIR / "output" / "checkpoints"
GENERATED_DIAGRAM_DIR = BASE_DIR / "cache" / "generated_diagrams"
ADVANCED_VISUAL_DIR   = BASE_DIR / "cache" / "advanced_visuals"

# ─────────────────────────────────────────────────────────────────────────────
# Document style requirements
# ─────────────────────────────────────────────────────────────────────────────

REQUIRED_HEADING_FONT    = "Arial"
REQUIRED_HEADING_SIZE_PT = 12
REQUIRED_HEADING_BOLD    = True

REQUIRED_BODY_FONT    = "Arial"
REQUIRED_BODY_SIZE_PT = 10

ALLOWED_HEADING_LEVELS = [1, 2, 3, 4, 5, 6]

# ─────────────────────────────────────────────────────────────────────────────
# OCR
# ─────────────────────────────────────────────────────────────────────────────

MIN_TEXT_CHARS_FOR_USEFUL_OCR = 15
MIN_CODE_HINT_SCORE           = 2
MIN_AVG_OCR_CONFIDENCE        = 0.35

# ─────────────────────────────────────────────────────────────────────────────
# Review comment metadata
# ─────────────────────────────────────────────────────────────────────────────

COMMENT_AUTHOR   = "Courseware Review System"
COMMENT_INITIALS = "CRS"

CAPTION_SEARCH_WINDOW = 6
SOURCE_SEARCH_WINDOW  = 7

# ─────────────────────────────────────────────────────────────────────────────
# LLM rewrite  (Step 3)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_LLM_REWRITE          = True
MISTRAL_REWRITE_MODEL       = "mistral-small-latest"
MAX_REWRITE_PARAGRAPHS      = 40
REWRITE_MIN_PARAGRAPH_CHARS = 60

# ─────────────────────────────────────────────────────────────────────────────
# Winston AI plagiarism scan  (Step 5)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_WINSTON               = True  # set False to skip and save credits during testing
WINSTON_API_URL              = "https://api.gowinston.ai/v2/plagiarism"
WINSTON_LANGUAGE             = "auto"
WINSTON_COUNTRY              = "us"
WINSTON_MIN_TEXT_CHARS       = 100
WINSTON_MAX_TEXT_CHARS       = 120_000
WINSTON_EXCLUDED_SOURCES     = ["ibm.com"]
WINSTON_SCORE_FLAG_THRESHOLD = 15
WINSTON_HIGH_SCORE_THRESHOLD = 35
WINSTON_ONLY_SUSPICIOUS      = True
WINSTON_FALLBACK_SCAN_IF_FEW = True
WINSTON_MIN_CANDIDATES       = 5
WINSTON_MAX_PARAGRAPHS       = 20

# ─────────────────────────────────────────────────────────────────────────────
# Accuracy checking  (Step 7)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_ACCURACY_CHECK          = True
ACCURACY_MODEL                 = "mistral-small-latest"
MAX_UNITS_FOR_ACCURACY         = 20
ACCURACY_MIN_UNIT_TEXT_CHARS   = 200
MAX_ACCURACY_FINDINGS_PER_UNIT = 5

# ─────────────────────────────────────────────────────────────────────────────
# Quiz generation  (Step 7)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_QUIZ_GENERATION      = True
QUIZ_MODEL                  = "mistral-small-latest"
MAX_UNITS_FOR_QUIZ          = 25
UNIT_HEADING_LEVEL_FOR_QUIZ = 1
CHAPTER_KEYWORDS            = ["chapter", "unit", "module", "lesson"]
QUIZ_MIN_UNIT_TEXT_CHARS    = 250

# ─────────────────────────────────────────────────────────────────────────────
# Checkpointing & pipeline integrity
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_CHECKPOINTS                = True
ENABLE_STAGE_POSTCONDITION_CHECKS = True  # validate state after each pipeline stage

# ─────────────────────────────────────────────────────────────────────────────
# Grammar checking
# ─────────────────────────────────────────────────────────────────────────────

GRAMMAR_MIN_PARAGRAPH_CHARS = 20    # skip paragraphs shorter than this
GRAMMAR_MAX_PARAGRAPH_CHARS = 5000  # skip extremely long paragraphs (e.g. tables dumped as text)

# ─────────────────────────────────────────────────────────────────────────────
# Duplicate detection
# ─────────────────────────────────────────────────────────────────────────────

MIN_PARAGRAPH_LEN_FOR_DUP_CHECK  = 80
NEAR_DUPLICATE_THRESHOLD         = 96
JACCARD_SIMILARITY_THRESHOLD     = 0.7
SEMANTIC_SIMILARITY_THRESHOLD    = 0.9

MAX_DUPLICATE_FINDINGS           = 200
MAX_NEAR_DUP_PER_PARAGRAPH       = 3
MAX_FUZZY_CANDIDATES             = 500   # cap input to fuzzy loop; limits O(n²) traversal
MAX_SEMANTIC_COMPARISONS         = 100   # cap input to semantic encoder; limits memory + encode time
FUZZY_CANDIDATE_WARNING_THRESHOLD = 300  # log a warning if candidates exceed this before fuzzy scan

ENABLE_SEMANTIC_DUPLICATES = True
EMBEDDING_MODEL_NAME       = "all-MiniLM-L6-v2"

# ─────────────────────────────────────────────────────────────────────────────
# Diagram recommendation  (Step 8)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_DIAGRAM_RECOMMENDATION = True
DIAGRAM_MIN_PARAGRAPH_CHARS   = 220
DIAGRAM_MAX_RECOMMENDATIONS   = 20
DIAGRAM_PROCESS_KEYWORDS = [
    "process", "workflow", "steps", "sequence", "flow", "lifecycle",
    "procedure", "execution", "pipeline", "architecture", "integration",
    "data flow", "movement", "routing", "stages", "phases",
    "compare", "comparison", "difference", "versus", "vs",
]

# ─────────────────────────────────────────────────────────────────────────────
# Diagram generation  (Step 9)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_DIAGRAM_GENERATION  = True
MAX_GENERATED_DIAGRAMS     = 10
DIAGRAM_IMAGE_WIDTH_INCHES = 5.5

# ─────────────────────────────────────────────────────────────────────────────
# Table & code analysis  (Step 10)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_TABLE_ANALYSIS = True
ENABLE_CODE_ANALYSIS  = True

TABLE_MIN_CELL_TEXT_CHARS   = 2
TABLE_REPEAT_ROW_THRESHOLD  = 0.95
TABLE_WIDE_COLUMN_THRESHOLD = 6

CODE_MIN_SYMBOL_SCORE  = 3
CODE_MIN_LINE_INDENT   = 2
CODE_BLOCK_STYLE_NAMES = ["code", "source code", "preformatted"]

# ─────────────────────────────────────────────────────────────────────────────
# Advanced visual generation  (Step 11)
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_ADVANCED_VISUALS            = True
MAX_ADVANCED_VISUALS               = 12
ADVANCED_VISUAL_WIDTH_INCHES       = 5.5
VISUAL_ENTITY_LIMIT                = 8
VISUAL_RELATION_LIMIT              = 10
ENABLE_LLM_VISUAL_SPEC_ENHANCEMENT = False  # keep deterministic by default

# ─────────────────────────────────────────────────────────────────────────────
# Ollama fallback  (used when MISTRAL_API_KEY is not set)
# Install: https://ollama.com  →  ollama pull llama3
# ─────────────────────────────────────────────────────────────────────────────

ENABLE_OLLAMA_FALLBACK = True
OLLAMA_BASE_URL        = "http://localhost:11434"
OLLAMA_REWRITE_MODEL   = "llama3"
OLLAMA_QUIZ_MODEL      = "llama3"
OLLAMA_ACCURACY_MODEL  = "llama3"
