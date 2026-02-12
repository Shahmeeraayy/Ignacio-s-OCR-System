from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd


SHEET_COLUMNS: dict[str, list[str]] = {
    "document_metadata": [
        "file",
        "path",
        "pages",
        "creator",
        "producer",
        "creation_date",
        "is_encrypted",
        "parse_timestamp",
    ],
    "pages": [
        "file",
        "page",
        "width",
        "height",
        "rotation",
        "text_chars",
        "word_count",
        "line_count",
        "table_count",
        "image_count",
        "link_count",
        "used_ocr",
    ],
    "text_lines": [
        "file",
        "page",
        "line_index",
        "x0",
        "top",
        "x1",
        "bottom",
        "text",
    ],
    "text_words": [
        "file",
        "page",
        "word_index",
        "x0",
        "top",
        "x1",
        "bottom",
        "text",
        "source",
    ],
    "text_chars": [
        "file",
        "page",
        "char_index",
        "x0",
        "top",
        "x1",
        "bottom",
        "text",
        "fontname",
        "size",
        "source",
    ],
    "tables_raw": [
        "file",
        "page",
        "table_index",
        "row_index",
        "col_index",
        "cell_text",
    ],
    "line_items_parsed": [
        "file",
        "item_index",
        "service_name",
        "sku",
        "units_qty",
        "term_start",
        "term_end",
        "list_unit_price_raw",
        "discount_pct_raw",
        "net_unit_price_raw",
        "net_total_raw",
        "description_continuation",
        "list_unit_price_value",
        "discount_pct_value",
        "net_unit_price_value",
        "net_total_value",
    ],
    "links": [
        "file",
        "page",
        "link_index",
        "uri",
        "rect_x0",
        "rect_y0",
        "rect_x1",
        "rect_y1",
    ],
    "images": [
        "file",
        "page",
        "image_index",
        "x0",
        "top",
        "x1",
        "bottom",
        "width",
        "height",
        "bits",
        "colorspace",
    ],
    "business_summary": [
        "file",
        "quote_number",
        "expiration_date",
        "subscription_period",
        "payment_method",
        "total_raw",
        "total_value",
        "currency",
        "overall_total_raw",
        "overall_total_value",
        "payment_year_1_raw",
        "payment_year_2_raw",
        "payment_year_3_raw",
        "regional_director",
        "regional_director_email",
        "payment_terms",
        "line_items_total_value",
        "expiration_date_iso",
        "error",
    ],
    "validation_report": [
        "file",
        "rule_id",
        "severity",
        "status",
        "observed_value",
        "expected_value",
        "details",
    ],
}


SORT_KEYS: dict[str, list[str]] = {
    "document_metadata": ["file"],
    "pages": ["file", "page"],
    "text_lines": ["file", "page", "line_index"],
    "text_words": ["file", "page", "word_index"],
    "text_chars": ["file", "page", "char_index"],
    "tables_raw": ["file", "page", "table_index", "row_index", "col_index"],
    "line_items_parsed": ["file", "item_index"],
    "links": ["file", "page", "link_index"],
    "images": ["file", "page", "image_index"],
    "business_summary": ["file"],
    "validation_report": ["file", "rule_id"],
}


def _rows_to_dataframe(sheet_name: str, rows: list[dict[str, Any]]) -> pd.DataFrame:
    columns = SHEET_COLUMNS[sheet_name]
    if not rows:
        return pd.DataFrame(columns=columns)

    normalized_rows: list[dict[str, Any]] = []
    for row in rows:
        normalized = {col: row.get(col) for col in columns}
        normalized_rows.append(normalized)

    frame = pd.DataFrame(normalized_rows)
    sort_by = [key for key in SORT_KEYS.get(sheet_name, []) if key in frame.columns]
    if sort_by:
        frame = frame.sort_values(sort_by, kind="mergesort")
    return frame


def write_excel(
    output_path: Path,
    rows_by_sheet: dict[str, list[dict[str, Any]]],
    include_char_layer: bool,
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        ordered_sheets = [
            "document_metadata",
            "pages",
            "text_lines",
            "text_words",
            "tables_raw",
            "line_items_parsed",
            "links",
            "images",
            "business_summary",
            "validation_report",
        ]
        if include_char_layer:
            ordered_sheets.insert(4, "text_chars")

        for sheet in ordered_sheets:
            frame = _rows_to_dataframe(sheet, rows_by_sheet.get(sheet, []))
            frame.to_excel(writer, sheet_name=sheet, index=False)


def write_json(json_path: Path, payload: dict[str, Any]) -> None:
    with json_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, ensure_ascii=False)

