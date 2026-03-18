import re


def clean_ocr_text(text: str) -> str:
    if not text:
        return ""

    text = text.replace(" ,", ",")
    text = text.replace(" .", ".")
    text = text.replace(" :", ":")
    text = text.replace(" ;", ";")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def merge_lines_to_text(lines):
    if not lines:
        return ""
    return "\n".join(line for line in lines if line.strip())
