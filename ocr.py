"""OCR: turn a receipt photo into text using Tesseract."""

import os
import re
import shutil

import pytesseract
from PIL import Image, ImageFilter, ImageOps

# Locate the tesseract binary: env var > PATH > default Windows install dir.
_DEFAULT_WIN = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_cmd = os.environ.get("TESSERACT_PATH") or shutil.which("tesseract") or _DEFAULT_WIN
pytesseract.pytesseract.tesseract_cmd = _cmd

_MONEY = re.compile(r"\d{1,5}[.,]\d{2}\b")


def _preprocess(img: Image.Image) -> Image.Image:
    """Clean up a phone photo so Tesseract has an easier time."""
    img = ImageOps.exif_transpose(img)
    img = img.convert("L")  # grayscale
    # Upscale small photos; receipts need ~300dpi-equivalent for good OCR.
    if img.width < 1500:
        scale = 1500 / img.width
        img = img.resize((1500, int(img.height * scale)), Image.LANCZOS)
    img = ImageOps.autocontrast(img, cutoff=2)
    img = img.filter(ImageFilter.SHARPEN)
    return img


def extract_text(image_path: str) -> str:
    """OCR the image. Tries two page-segmentation modes and keeps the one
    that produced more money-looking amounts (a decent proxy for receipt
    OCR quality)."""
    img = _preprocess(Image.open(image_path))

    best_text, best_score = "", -1
    for psm in (4, 6):  # 4 = sparse columns, 6 = uniform block
        try:
            text = pytesseract.image_to_string(img, config=f"--oem 3 --psm {psm}")
        except pytesseract.TesseractError:
            continue
        score = len(_MONEY.findall(text))
        if score > best_score:
            best_text, best_score = text, score
    return best_text
