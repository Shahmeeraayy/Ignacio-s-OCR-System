from __future__ import annotations

from pathlib import Path
from typing import Any


def _load_ocr_dependencies():
    try:
        from pdf2image import convert_from_path
    except ImportError as exc:
        raise RuntimeError("OCR mode needs `pdf2image`. Install it with pip.") from exc

    try:
        import pytesseract
        from pytesseract import Output
    except ImportError as exc:
        raise RuntimeError("OCR mode needs `pytesseract`. Install it with pip.") from exc

    return convert_from_path, pytesseract, Output


def ocr_page_words(
    pdf_path: Path,
    page_number: int,
    page_width: float,
    page_height: float,
    dpi: int,
    tesseract_cmd: str | None = None,
    poppler_path: str | None = None,
) -> tuple[str, list[dict[str, Any]]]:
    convert_from_path, pytesseract, output_type = _load_ocr_dependencies()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    images = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        first_page=page_number,
        last_page=page_number,
        poppler_path=poppler_path,
    )
    if not images:
        return "", []

    image = images[0]
    data = pytesseract.image_to_data(image, output_type=output_type.DICT)

    width_px, height_px = image.size
    x_scale = page_width / float(width_px)
    y_scale = page_height / float(height_px)

    words: list[dict[str, Any]] = []
    text_parts: list[str] = []
    count = len(data["text"])
    for idx in range(count):
        token = (data["text"][idx] or "").strip()
        if not token:
            continue
        left = float(data["left"][idx])
        top = float(data["top"][idx])
        width = float(data["width"][idx])
        height = float(data["height"][idx])
        x0 = left * x_scale
        y0 = top * y_scale
        x1 = (left + width) * x_scale
        y1 = (top + height) * y_scale
        words.append(
            {
                "text": token,
                "x0": x0,
                "top": y0,
                "x1": x1,
                "bottom": y1,
                "source": "ocr",
            }
        )
        text_parts.append(token)

    return " ".join(text_parts), words

