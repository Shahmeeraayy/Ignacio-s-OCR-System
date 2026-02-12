from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pdfplumber
from pypdf import PdfReader

from .ocr import ocr_page_words


def _stringify_colorspace(colorspace: Any) -> str | None:
    if colorspace is None:
        return None
    if isinstance(colorspace, (list, tuple)):
        return ",".join(str(part) for part in colorspace)
    return str(colorspace)


def _extract_links(reader: PdfReader, page_number: int) -> list[dict[str, Any]]:
    page = reader.pages[page_number - 1]
    annots = page.get("/Annots")
    links: list[dict[str, Any]] = []
    if not annots:
        return links

    link_index = 1
    for annot_ref in annots:
        try:
            annot = annot_ref.get_object()
        except Exception:
            continue
        subtype = str(annot.get("/Subtype"))
        if subtype != "/Link":
            continue

        action = annot.get("/A")
        uri = None
        if action and action.get("/URI"):
            uri = str(action.get("/URI"))

        rect = annot.get("/Rect")
        rect_values = [None, None, None, None]
        if rect and len(rect) == 4:
            rect_values = [float(value) for value in rect]

        links.append(
            {
                "page": page_number,
                "link_index": link_index,
                "uri": uri,
                "rect_x0": rect_values[0],
                "rect_y0": rect_values[1],
                "rect_x1": rect_values[2],
                "rect_y1": rect_values[3],
            }
        )
        link_index += 1

    return links


def _build_lines_from_words(words: list[dict[str, Any]], y_tolerance: float = 2.5) -> list[dict[str, Any]]:
    if not words:
        return []
    ordered = sorted(words, key=lambda item: (round(float(item["top"]), 3), float(item["x0"])))

    line_groups: list[list[dict[str, Any]]] = []
    for word in ordered:
        if not line_groups:
            line_groups.append([word])
            continue
        reference_top = line_groups[-1][0]["top"]
        if abs(float(word["top"]) - float(reference_top)) <= y_tolerance:
            line_groups[-1].append(word)
        else:
            line_groups.append([word])

    lines: list[dict[str, Any]] = []
    for idx, group in enumerate(line_groups, start=1):
        sorted_group = sorted(group, key=lambda item: float(item["x0"]))
        lines.append(
            {
                "line_index": idx,
                "x0": min(float(item["x0"]) for item in sorted_group),
                "top": min(float(item["top"]) for item in sorted_group),
                "x1": max(float(item["x1"]) for item in sorted_group),
                "bottom": max(float(item["bottom"]) for item in sorted_group),
                "text": " ".join(str(item["text"]) for item in sorted_group),
            }
        )
    return lines


def _normalize_word(word: dict[str, Any], source: str) -> dict[str, Any]:
    return {
        "text": str(word.get("text", "")),
        "x0": float(word.get("x0", 0.0)),
        "top": float(word.get("top", 0.0)),
        "x1": float(word.get("x1", 0.0)),
        "bottom": float(word.get("bottom", 0.0)),
        "source": source,
    }


def extract_pdf_raw(
    pdf_path: Path,
    config: dict[str, Any],
    ocr_mode: str = "auto",
    include_char_layer: bool = False,
    include_tables: bool = True,
    tesseract_cmd: str | None = None,
    poppler_path: str | None = None,
) -> dict[str, Any]:
    reader = PdfReader(str(pdf_path))
    metadata = reader.metadata or {}
    parse_timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    meta_row = {
        "file": pdf_path.name,
        "path": str(pdf_path),
        "pages": len(reader.pages),
        "creator": metadata.get("/Creator"),
        "producer": metadata.get("/Producer"),
        "creation_date": metadata.get("/CreationDate"),
        "is_encrypted": bool(reader.is_encrypted),
        "parse_timestamp": parse_timestamp,
    }

    results: dict[str, Any] = {
        "metadata": meta_row,
        "pages": [],
        "text_lines": [],
        "text_words": [],
        "text_chars": [],
        "tables_raw": [],
        "tables_structured": [],
        "links": [],
        "images": [],
        "full_text": "",
        "ocr_pages": [],
        "error": None,
    }

    ocr_min_chars = int(config.get("ocr", {}).get("min_native_text_chars", 50))
    table_settings = config.get("table_settings", {})

    with pdfplumber.open(str(pdf_path)) as pdf:
        full_text_parts: list[str] = []
        for page_idx, page in enumerate(pdf.pages, start=1):
            native_text = (page.extract_text() or "").strip()
            native_words = [_normalize_word(word, "native") for word in (page.extract_words() or [])]

            selected_text = native_text
            selected_words = native_words
            page_used_ocr = False

            should_ocr = ocr_mode == "always" or (ocr_mode == "auto" and len(native_text) < ocr_min_chars)
            if should_ocr:
                try:
                    ocr_text, ocr_words = ocr_page_words(
                        pdf_path=pdf_path,
                        page_number=page_idx,
                        page_width=float(page.width),
                        page_height=float(page.height),
                        dpi=int(config.get("ocr", {}).get("dpi", 300)),
                        tesseract_cmd=tesseract_cmd,
                        poppler_path=poppler_path,
                    )
                    if ocr_words:
                        selected_words = ocr_words
                        selected_text = ocr_text
                        page_used_ocr = True
                except Exception:
                    # OCR is best-effort, native extraction remains available.
                    page_used_ocr = False

            page_lines = _build_lines_from_words(selected_words)
            full_text_parts.append(selected_text)

            links = _extract_links(reader, page_idx)
            for link in links:
                link["file"] = pdf_path.name
                results["links"].append(link)

            for image_idx, image in enumerate(page.images, start=1):
                results["images"].append(
                    {
                        "file": pdf_path.name,
                        "page": page_idx,
                        "image_index": image_idx,
                        "x0": float(image.get("x0", 0.0)),
                        "top": float(image.get("top", 0.0)),
                        "x1": float(image.get("x1", 0.0)),
                        "bottom": float(image.get("bottom", 0.0)),
                        "width": float(image.get("width", 0.0)),
                        "height": float(image.get("height", 0.0)),
                        "bits": image.get("bits"),
                        "colorspace": _stringify_colorspace(image.get("colorspace")),
                    }
                )

            if include_char_layer and not page_used_ocr:
                for char_idx, char in enumerate(page.chars, start=1):
                    results["text_chars"].append(
                        {
                            "file": pdf_path.name,
                            "page": page_idx,
                            "char_index": char_idx,
                            "x0": float(char.get("x0", 0.0)),
                            "top": float(char.get("top", 0.0)),
                            "x1": float(char.get("x1", 0.0)),
                            "bottom": float(char.get("bottom", 0.0)),
                            "text": str(char.get("text", "")),
                            "fontname": char.get("fontname"),
                            "size": char.get("size"),
                            "source": "native",
                        }
                    )

            for word_idx, word in enumerate(selected_words, start=1):
                results["text_words"].append(
                    {
                        "file": pdf_path.name,
                        "page": page_idx,
                        "word_index": word_idx,
                        "x0": float(word["x0"]),
                        "top": float(word["top"]),
                        "x1": float(word["x1"]),
                        "bottom": float(word["bottom"]),
                        "text": str(word["text"]),
                        "source": word.get("source", "native"),
                    }
                )

            for line in page_lines:
                results["text_lines"].append(
                    {
                        "file": pdf_path.name,
                        "page": page_idx,
                        "line_index": line["line_index"],
                        "x0": line["x0"],
                        "top": line["top"],
                        "x1": line["x1"],
                        "bottom": line["bottom"],
                        "text": line["text"],
                    }
                )

            page_tables = []
            if include_tables:
                try:
                    page_tables = page.extract_tables(table_settings or None) or []
                except Exception:
                    page_tables = []
                for table_idx, table in enumerate(page_tables, start=1):
                    rows = table or []
                    results["tables_structured"].append(
                        {
                            "file": pdf_path.name,
                            "page": page_idx,
                            "table_index": table_idx,
                            "rows": rows,
                        }
                    )
                    for row_idx, row in enumerate(rows, start=1):
                        normalized_row = row if isinstance(row, list) else [row]
                        for col_idx, cell in enumerate(normalized_row, start=1):
                            results["tables_raw"].append(
                                {
                                    "file": pdf_path.name,
                                    "page": page_idx,
                                    "table_index": table_idx,
                                    "row_index": row_idx,
                                    "col_index": col_idx,
                                    "cell_text": cell,
                                }
                            )

            page_row = {
                "file": pdf_path.name,
                "page": page_idx,
                "width": float(page.width),
                "height": float(page.height),
                "rotation": int(page.rotation or 0),
                "text_chars": len(selected_text),
                "word_count": len(selected_words),
                "line_count": len(page_lines),
                "table_count": len(page_tables),
                "image_count": len(page.images),
                "link_count": len(links),
                "used_ocr": page_used_ocr,
            }
            results["pages"].append(page_row)
            if page_used_ocr:
                results["ocr_pages"].append(page_idx)

    results["full_text"] = "\n".join(part for part in full_text_parts if part)
    return results

