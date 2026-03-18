import re


CODE_KEYWORDS = {
    "def", "class", "import", "return", "if", "else", "elif", "for", "while",
    "try", "except", "public", "private", "static", "void", "int", "string",
    "const", "let", "var", "function", "select", "from", "where", "join",
    "{", "}", "(", ")", ";", "==", "!=", ">=", "<=", "=>", "::", "#include"
}


def score_code_like_text(text: str) -> int:
    if not text:
        return 0

    score = 0
    lower = text.lower()

    for kw in CODE_KEYWORDS:
        if kw in lower:
            score += 1

    if re.search(r"^[ \t]{2,}\S+", text, flags=re.MULTILINE):
        score += 1
    if re.search(r"[{}();=<>\[\]#]", text):
        score += 1
    if re.search(r"\b[a-zA-Z_][a-zA-Z0-9_]*\s*\(", text):
        score += 1

    return score


def classify_by_ocr_preview(preview_text: str) -> str:
    if not preview_text or not preview_text.strip():
        return "unknown"

    score = score_code_like_text(preview_text)
    alpha_chars = sum(ch.isalpha() for ch in preview_text)
    special_chars = sum(ch in "{}[]()#;=<>:/\\._-" for ch in preview_text)

    if score >= 2 or (special_chars > 10 and alpha_chars > 10):
        return "code"

    line_count = len([line for line in preview_text.splitlines() if line.strip()])
    if line_count >= 2:
        return "text"

    return "mixed"


def fallback_classify_from_filename(filename: str) -> str:
    lower = filename.lower()
    if "code" in lower or "snippet" in lower or "terminal" in lower:
        return "code"
    if "diagram" in lower or "flow" in lower:
        return "mixed"
    return "unknown"
