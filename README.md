# Courseware Review System

A modular DOCX review-and-fix pipeline for courseware books, modules, and training manuals.

It produces exactly two Word outputs for each input file:
- `<Actual File Name>_Review Comments.docx`
- `<Actual File Name>_Final Fixed.docx`

---

## What the pipeline does

The system combines deterministic checks with targeted API-assisted steps.

### Deterministic-first checks
- Heading and font validation using the document's own dominant font — does not hardcode Arial
- Navigation-pane and heading hierarchy checks
- Spacing cleanup (in-place, preserves bold/italic/hyperlinks)
- Duplicate paragraph detection with structural repeat suppression (repeating page headers flagged once, not per occurrence)
- Near-duplicate detection with Jaccard pre-filter for performance on large documents
- Semantic duplicate detection via sentence-transformers (batched, full document)
- Image extraction, figure number and source-link checks
- Table-aware and code-block-aware checks
- Chapter/unit detection for quiz placement
- Advanced visual classification and rendering
- TOC entries excluded from font checks — never flagged as inconsistent

### API-assisted steps (Mistral or Ollama fallback)
- Winston AI for plagiarism/originality screening
- Mistral (or Ollama) for selective paragraph rewriting
- Mistral (or Ollama) for quiz generation
- Mistral (or Ollama) for content accuracy checking

---

## Key behavior

### Dominant font detection
The pipeline detects the document's own dominant heading and body fonts before running any checks or fixes. It never forces Arial onto a document that uses a different font consistently — it only flags or corrects genuine outliers.

### Ollama fallback
If `MISTRAL_API_KEY` is not set, the pipeline automatically falls back to a locally running Ollama instance. Install Ollama from https://ollama.com and run:

```bash
ollama pull llama3
```

Priority order:
1. `MISTRAL_API_KEY` present → Mistral (fast, hosted)
2. Key missing + Ollama running → Ollama (free, local, slower)
3. Neither → LLM stages skip silently

### Unit-level quiz generation only
Quizzes are generated after each **unit/chapter**, not after each topic.

- `UNIT_HEADING_LEVEL_FOR_QUIZ=1` means only `Heading 1` is treated as a chapter/unit boundary
- Chapter detection also prefers headings containing words like `chapter`, `unit`, `module`, or `lesson`
- Quiz insertion is drift-safe: uses heading text match in the live document rather than pre-recorded paragraph indices that shift after rewrites

### Final Fixed document is comment-free
The Final Fixed output strips all existing comments (including any from previous manual reviewers) before applying fixes. Comments belong only in the Review Comments document.

### Review Comments document
Includes a summary page prepended at the top showing:
- Total issues found
- Severity breakdown (High / Medium / Low / Info)
- Top issue types with counts
- Units detected with quiz status

### Selective Winston scanning
Winston scanning is intentionally restricted to paragraphs that are long enough, not headings, not duplicates, and suspicious enough to justify the cost.

### Pipeline resilience
Each pipeline stage uses lightweight field-level backup instead of deep-copying the entire state. If a stage fails, only the mutable fields are rolled back — ingest data is never lost. The UI shows which stages were skipped and why.

### Checkpointing
The pipeline saves checkpoint JSON files after major stages so large documents can resume after interruption without rerunning completed stages.

---

## Project dependencies

### Python packages
Install with:

```bash
pip install -r requirements.txt
```

### System packages
Required in Colab, Ubuntu, or Debian-based environments:

```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr graphviz
```

---

## Environment variables

Create a `.env` file in the project root.

```env
MISTRAL_API_KEY=your_mistral_api_key_here
WINSTON_API_KEY=your_winston_api_key_here
SERPAPI_KEY=your_serp_api_key_here
```

All three are optional if Ollama is running locally — LLM features will use Ollama, plagiarism scan and image source lookup will be skipped with visible warnings in the UI.

---

## Recommended folder structure

```text
courseware_review_system/
├── input/raw_docx/
├── output/review_comments/
├── output/final_fixed/
├── output/intermediate_json/
├── output/checkpoints/
├── cache/extracted_images/
├── cache/ocr_results/
├── cache/generated_diagrams/
├── cache/advanced_visuals/
└── src/
```

---

## Main pipeline stages

### Step 1
- DOCX ingestion with dominant font detection
- Heading heuristic fallback for documents without Word heading styles
- Formatting checks against document's own dominant style
- Heading hierarchy and navigation pane checks
- Spacing checks
- Image extraction

### Step 2
- OCR for image text and code screenshots
- Image OCR comments

### Step 3
- Figure number detection
- Image source-link detection
- Nearby caption analysis

### Step 4
- Exact duplicate detection (structural repeats collapsed to one finding)
- Near-duplicate detection with Jaccard pre-filter
- Semantic duplicate detection (full document, batched encoding)
- Repeated heading detection (structural headings excluded)

### Step 5
- Winston AI plagiarism/originality screening

### Step 6
- Mistral (or Ollama) rewrite for flagged sections only

### Step 7
- Quiz detection and generation after each chapter/unit

### Step 8
- Diagram recommendation and quiz insertion after units (drift-safe)

### Step 9
- Deterministic diagram generation and insertion

### Step 10
- Table-aware and code-aware analysis
- Final polish

### Step 11
- Advanced visual classification
- Architecture/deployment/sequence/concept visual generation
- Final issue consolidation with section-level grouping for format issues

---

## Typical run

### Streamlit UI (recommended)
```bash
streamlit run app.py
```

Upload any `.docx` file and click **Run Review Pipeline**. The UI shows:
- Live pipeline progress
- Feature status banners (Mistral/Ollama/Winston/SerpAPI)
- Issue summary with severity breakdown
- Filterable issue list by severity and type
- Download buttons for both output documents
- Units detected with quiz status

### CLI (individual stages)
```bash
python run_step11.py
```

All stage runners resume automatically from the latest checkpoint for the same input filename.

---

## Output contract

For each input DOCX the pipeline outputs:
- `*_Review Comments.docx` — original-like file with inline Word comments, issue highlights, and a summary page
- `*_Final Fixed.docx` — corrected file with normalized formatting, inserted quizzes, visuals, safe fixes, and zero comments

---

## Notes on requirements.txt

`requirements.txt` covers Python packages only. Two required system packages are not pip-installed:
- `tesseract-ocr`
- `graphviz`

---

## Troubleshooting

### `ModuleNotFoundError: No module named ...`
```bash
pip install -r requirements.txt
```

### `ModuleNotFoundError: No module named 'sentence_transformers'`
```bash
pip install sentence-transformers
```

### Tesseract OCR not found
```bash
sudo apt-get install -y tesseract-ocr
```

### Graphviz rendering fails
```bash
sudo apt-get install -y graphviz
```

### Ollama not reachable
Install from https://ollama.com, then:
```bash
ollama pull llama3
ollama serve
```

### Winston API issues
- Check `WINSTON_API_KEY` is present in `.env`
- Check request limits and credits
- Verify `ENABLE_WINSTON = True` in `config.py`

### Mistral API issues
- Check `MISTRAL_API_KEY` is present in `.env`
- Verify model name in config exists
- Check rewrite/quiz/accuracy flags are enabled in `config.py`



