from __future__ import annotations

import re
from typing import Any

from .normalize import parse_currency_code, parse_currency_value, parse_date_value, split_term_range


def _clean_cell(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def _first_match(text: str, patterns: list[str]) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            if match.groups():
                return match.group(1).strip()
            return match.group(0).strip()
    return None


def parse_line_items(
    file_name: str,
    tables_structured: list[dict[str, Any]],
    line_item_rules: dict[str, Any],
) -> list[dict[str, Any]]:
    expected_headers = [header.lower() for header in line_item_rules.get("header_contains", [])]
    min_columns = int(line_item_rules.get("min_columns", 8))

    items: list[dict[str, Any]] = []
    current_item: dict[str, Any] | None = None

    for table in tables_structured:
        rows = table.get("rows") or []
        for row in rows:
            if not isinstance(row, list):
                continue
            normalized = [_clean_cell(cell) for cell in row]
            if len(normalized) < min_columns:
                normalized.extend([None] * (min_columns - len(normalized)))

            row_text = " ".join(cell for cell in normalized if cell).lower()
            if expected_headers and all(header in row_text for header in expected_headers[:2]):
                continue
            if "service/product name" in row_text and "code/sku" in row_text:
                continue

            c1, c2, c3, c4, c5, c6, c7, c8 = normalized[:8]

            # "TOTAL" style rows usually only have the amount in final column.
            if not any([c1, c2, c3, c4, c5, c6, c7]) and c8:
                current_item = None
                continue

            is_continuation = bool(c1 and not any([c2, c3, c4, c5, c6, c7, c8]))
            if is_continuation and current_item:
                existing = current_item.get("description_continuation")
                current_item["description_continuation"] = (
                    f"{existing}\n{c1}".strip() if existing else c1
                )
                continue

            # Primary detail rows are expected to include SKU and pricing columns.
            if c2 and any([c5, c7, c8]):
                term_start, term_end = split_term_range(c4)
                item = {
                    "file": file_name,
                    "item_index": len(items) + 1,
                    "service_name": c1,
                    "sku": c2,
                    "units_qty": c3,
                    "term_start": term_start,
                    "term_end": term_end,
                    "list_unit_price_raw": c5,
                    "discount_pct_raw": c6,
                    "net_unit_price_raw": c7,
                    "net_total_raw": c8,
                    "description_continuation": None,
                    "list_unit_price_value": parse_currency_value(c5),
                    "discount_pct_value": _parse_percent(c6),
                    "net_unit_price_value": parse_currency_value(c7),
                    "net_total_value": parse_currency_value(c8),
                }
                items.append(item)
                current_item = item
                continue

            # Any text-only row after an item is treated as continuation detail.
            if c1 and current_item:
                existing = current_item.get("description_continuation")
                current_item["description_continuation"] = (
                    f"{existing}\n{c1}".strip() if existing else c1
                )

    return items


def _parse_percent(raw: str | None) -> float | None:
    if not raw:
        return None
    match = re.search(r"[-+]?[0-9]+(?:\.[0-9]+)?", raw)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def _extract_director_fields(
    tables_structured: list[dict[str, Any]],
) -> tuple[str | None, str | None, str | None]:
    for table in tables_structured:
        rows = table.get("rows") or []
        for idx, row in enumerate(rows):
            if not isinstance(row, list):
                continue
            cells = [(_clean_cell(cell) or "") for cell in row]
            row_text = " ".join(cells).lower()
            if "regional director" in row_text and "payment terms" in row_text:
                if idx + 1 >= len(rows):
                    continue
                next_row = rows[idx + 1]
                if isinstance(next_row, list) and len(next_row) >= 4:
                    director = _clean_cell(next_row[0])
                    email = _clean_cell(next_row[1])
                    payment_terms = _clean_cell(next_row[3])
                    return director, email, payment_terms
    return None, None, None


def extract_business_summary(
    file_name: str,
    full_text: str,
    tables_structured: list[dict[str, Any]],
    line_items: list[dict[str, Any]],
    config: dict[str, Any],
) -> dict[str, Any]:
    patterns = config.get("field_patterns", {})
    normalization = config.get("normalization", {})
    date_formats = normalization.get("date_input_formats", ["%m/%d/%Y"])
    currency_default = normalization.get("currency_default", "USD")

    summary: dict[str, Any] = {"file": file_name}
    for field_name, field_patterns in patterns.items():
        if isinstance(field_patterns, list):
            summary[field_name] = _first_match(full_text, field_patterns)

    director, director_email, payment_terms = _extract_director_fields(tables_structured)
    if director:
        summary["regional_director"] = director
    if director_email:
        summary["regional_director_email"] = director_email
    if payment_terms:
        summary["payment_terms"] = payment_terms

    summary["expiration_date_iso"] = parse_date_value(summary.get("expiration_date"), date_formats)

    total_raw = summary.get("total_raw")
    overall_total_raw = summary.get("overall_total_raw")
    summary["total_value"] = parse_currency_value(total_raw)
    summary["overall_total_value"] = parse_currency_value(overall_total_raw)
    summary["currency"] = parse_currency_code(total_raw or overall_total_raw, currency_default)

    line_total = sum(item.get("net_total_value") or 0.0 for item in line_items)
    summary["line_items_total_value"] = line_total if line_items else None
    summary["error"] = None
    return summary
