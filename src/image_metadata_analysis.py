import re

from config import CAPTION_SEARCH_WINDOW, SOURCE_SEARCH_WINDOW


FIGURE_PATTERNS = [
    r'^\s*figure\s*\d+[\.: -]?',
    r'^\s*fig\.\s*\d+[\.: -]?',
    r'^\s*fig\s*\d+[\.: -]?',
]

SOURCE_PATTERNS = [
    r'https?://\S+',
    r'www\.\S+',
    r'\bsource\s*:\s*.+',
    r'\breference\s*:\s*.+',
    r'\bcourtesy\s*:\s*.+',
]


def _matches_any(text: str, patterns):
    if not text:
        return False
    lower = text.strip().lower()
    return any(re.search(pattern, lower, flags=re.IGNORECASE) for pattern in patterns)


def is_caption_like(text: str) -> bool:
    if not text or not text.strip():
        return False

    # Must explicitly match a figure/caption pattern
    if _matches_any(text, FIGURE_PATTERNS):
        return True

    return False


def extract_figure_number(text: str):
    if not text:
        return None

    patterns = [
        r'(figure\s*\d+)',
        r'(fig\.\s*\d+)',
        r'(fig\s*\d+)',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            return m.group(1)
    return None


def extract_source_link_text(text: str):
    if not text:
        return None

    url_match = re.search(r'(https?://\S+|www\.\S+)', text, flags=re.IGNORECASE)
    if url_match:
        return url_match.group(1)

    if re.search(r'\b(source|reference|courtesy)\s*:', text, flags=re.IGNORECASE):
        return text.strip()

    return None


def analyze_image_neighbors(state):
    for img in state.images:
        base_idx = img.paragraph_index
        if base_idx is None or not state.paragraphs:
            continue

        caption_candidate_idx = None
        caption_candidate_text = None
        figure_number_text = None
        source_link_text = None

        start_idx = max(0, base_idx - CAPTION_SEARCH_WINDOW)
        end_idx = min(len(state.paragraphs) - 1, base_idx + CAPTION_SEARCH_WINDOW)

        # Caption / figure-number search — never treat a heading as a caption
        for idx in range(start_idx, end_idx + 1):
            para = state.paragraphs[idx]
            if para.is_heading:
                continue
            para_text = para.text
            if not para_text:
                continue

            if is_caption_like(para_text):
                caption_candidate_idx = idx
                caption_candidate_text = para_text
                fig_no = extract_figure_number(para_text)
                if fig_no:
                    figure_number_text = fig_no
                break

        # Source search — never treat a heading as a source
        source_start = max(0, base_idx - SOURCE_SEARCH_WINDOW)
        source_end = min(len(state.paragraphs) - 1, base_idx + SOURCE_SEARCH_WINDOW)
        for idx in range(source_start, source_end + 1):
            para = state.paragraphs[idx]
            if para.is_heading:
                continue
            para_text = para.text
            if not para_text:
                continue

            src = extract_source_link_text(para_text)
            if src:
                source_link_text = src
                if caption_candidate_idx is None:
                    caption_candidate_idx = idx
                    caption_candidate_text = para_text
                break

        img.nearest_caption_paragraph_index = caption_candidate_idx
        img.nearest_caption_text = caption_candidate_text
        img.has_figure_number = bool(figure_number_text)
        img.figure_number_text = figure_number_text
        img.has_source_link = bool(source_link_text)
        img.source_link_text = source_link_text

    return state