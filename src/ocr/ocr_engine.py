from pathlib import Path
from typing import List, Tuple
import platform

import cv2
import pytesseract

if platform.system() == "Windows":
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def preprocess_for_ocr(image_path: Path):
    img = cv2.imread(str(image_path))
    if img is None:
        return None

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.medianBlur(gray, 3)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return thresh


def run_ocr(image_path: Path) -> List[Tuple[str, float]]:
    processed = preprocess_for_ocr(image_path)
    if processed is None:
        return []

    data = pytesseract.image_to_data(
        processed,
        output_type=pytesseract.Output.DICT,
        config="--oem 3 --psm 6"
    )

    lines = []
    n = len(data["text"])

    for i in range(n):
        text = data["text"][i].strip()
        conf_raw = data["conf"][i]

        if not text:
            continue

        try:
            conf = float(conf_raw)
            if conf < 0:
                conf = 0.0
            conf = round(conf / 100.0, 4)
        except Exception:
            conf = 0.0

        lines.append((text, conf))

    return lines