import re


COMMON_REPLACEMENTS = [
    (r"\b0utput\b", "Output"),
    (r"\bretum\b", "return"),
    (r"\bRetum\b", "Return"),
    (r"\bimporf\b", "import"),
    (r"\bclas5\b", "class"),
    (r"\bwhi1e\b", "while"),
    (r"\be1se\b", "else"),
]


def clean_code_ocr_text(text: str) -> str:
    if not text:
        return ""

    cleaned = text

    for pattern, replacement in COMMON_REPLACEMENTS:
        cleaned = re.sub(pattern, replacement, cleaned)

    cleaned = cleaned.replace("“", '"').replace("”", '"')
    cleaned = cleaned.replace("‘", "'").replace("’", "'")
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)

    lines = cleaned.splitlines()
    normalized_lines = [line.rstrip() for line in lines]
    return "\n".join(normalized_lines).strip()


def looks_like_code(text: str) -> bool:
    if not text:
        return False

    patterns = [
        r"\bdef\s+\w+\s*\(",
        r"\bclass\s+\w+",
        r"\bimport\s+\w+",
        r"[{}();=]",
        r"\bif\b.*:",
        r"\bfor\b.*:",
        r"\bwhile\b.*:",
        r"#include",
    ]
    return any(re.search(p, text) for p in patterns)
