from __future__ import annotations

import re
from typing import Any

from .normalize import parse_currency_code, parse_currency_value, parse_date_value, split_term_range


LINE_ITEM_HEADER_PATTERNS: dict[str, tuple[str, ...]] = {
    "service_name": ("service product name", "product name"),
    "sku": ("service product code sku", "product code sku", "code sku", "product code", "sku"),
    "units_qty": ("units quantity", "quantity", "units"),
    "term": ("term",),
    "list_unit_price": ("list unit price",),
    "discount_pct": ("discount",),
    "net_unit_price": ("net unit price",),
    "net_total": ("net total",),
}


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


def _normalize_header_text(value: str | None) -> str:
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def _detect_line_item_column_map(cells: list[str | None]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(cells):
        normalized = _normalize_header_text(cell)
        if not normalized:
            continue
        for field_name, patterns in LINE_ITEM_HEADER_PATTERNS.items():
            if any(pattern in normalized for pattern in patterns):
                if field_name not in mapping:
                    mapping[field_name] = idx
                break

    has_identity = "service_name" in mapping and "sku" in mapping
    has_pricing = any(name in mapping for name in ("list_unit_price", "net_unit_price", "net_total"))
    if has_identity and has_pricing:
        return mapping
    return {}


def _mapped_cell(cells: list[str | None], mapping: dict[str, int], field_name: str) -> str | None:
    idx = mapping.get(field_name)
    if idx is None:
        return None
    if idx < 0 or idx >= len(cells):
        return None
    return cells[idx]


def parse_line_items(
    file_name: str,
    tables_structured: list[dict[str, Any]],
    line_item_rules: dict[str, Any],
) -> list[dict[str, Any]]:
    expected_headers = [header.lower() for header in line_item_rules.get("header_contains", [])]
    min_columns = int(line_item_rules.get("min_columns", 8))

    items: list[dict[str, Any]] = []
    current_item: dict[str, Any] | None = None
    quote_section_index = 1
    section_has_items = False
    active_column_map: dict[str, int] = {}

    for table in tables_structured:
        rows = table.get("rows") or []
        for row in rows:
            if not isinstance(row, list):
                continue
            normalized = [_clean_cell(cell) for cell in row]
            detected_map = _detect_line_item_column_map(normalized)
            if detected_map:
                active_column_map = detected_map
                continue

            row_text = " ".join(cell for cell in normalized if cell).lower()
            if expected_headers and all(header in row_text for header in expected_headers[:2]):
                continue
            if "service/product name" in row_text and ("code/sku" in row_text or "sku" in row_text):
                continue

            if active_column_map:
                service_name = _mapped_cell(normalized, active_column_map, "service_name")
                sku = _mapped_cell(normalized, active_column_map, "sku")
                units_qty = _mapped_cell(normalized, active_column_map, "units_qty")
                term_raw = _mapped_cell(normalized, active_column_map, "term")
                list_unit_price_raw = _mapped_cell(normalized, active_column_map, "list_unit_price")
                discount_pct_raw = _mapped_cell(normalized, active_column_map, "discount_pct")
                net_unit_price_raw = _mapped_cell(normalized, active_column_map, "net_unit_price")
                net_total_raw = _mapped_cell(normalized, active_column_map, "net_total")
            else:
                if len(normalized) < min_columns:
                    normalized.extend([None] * (min_columns - len(normalized)))
                c1, c2, c3, c4, c5, c6, c7, c8 = normalized[:8]
                service_name = c1
                sku = c2
                units_qty = c3
                term_raw = c4
                list_unit_price_raw = c5
                discount_pct_raw = c6
                net_unit_price_raw = c7
                net_total_raw = c8

            # "TOTAL" style rows usually only have the amount in final column.
            if not any(
                [service_name, sku, units_qty, term_raw, list_unit_price_raw, discount_pct_raw, net_unit_price_raw]
            ) and net_total_raw:
                current_item = None
                if section_has_items:
                    quote_section_index += 1
                    section_has_items = False
                continue

            non_empty_cells = [cell for cell in normalized if cell]
            if current_item and len(non_empty_cells) == 1:
                text = non_empty_cells[0]
                existing = current_item.get("description_continuation")
                current_item["description_continuation"] = (
                    f"{existing}\n{text}".strip() if existing else text
                )
                continue

            is_continuation = bool(
                service_name
                and not any(
                    [sku, units_qty, term_raw, list_unit_price_raw, discount_pct_raw, net_unit_price_raw, net_total_raw]
                )
            )
            if is_continuation and current_item:
                existing = current_item.get("description_continuation")
                current_item["description_continuation"] = (
                    f"{existing}\n{service_name}".strip() if existing else service_name
                )
                continue

            # Primary detail rows are expected to include SKU and pricing columns.
            if sku and any([list_unit_price_raw, net_unit_price_raw, net_total_raw]):
                term_start, term_end = split_term_range(term_raw)
                item = {
                    "file": file_name,
                    "item_index": len(items) + 1,
                    "service_name": service_name,
                    "sku": sku,
                    "units_qty": units_qty,
                    "term_start": term_start,
                    "term_end": term_end,
                    "list_unit_price_raw": list_unit_price_raw,
                    "discount_pct_raw": discount_pct_raw,
                    "net_unit_price_raw": net_unit_price_raw,
                    "net_total_raw": net_total_raw,
                    "description_continuation": None,
                    "list_unit_price_value": parse_currency_value(list_unit_price_raw),
                    "discount_pct_value": _parse_percent(discount_pct_raw),
                    "net_unit_price_value": parse_currency_value(net_unit_price_raw),
                    "net_total_value": parse_currency_value(net_total_raw),
                    "quote_section_index": quote_section_index,
                }
                items.append(item)
                current_item = item
                section_has_items = True
                continue

            # Any text-only row after an item is treated as continuation detail.
            if service_name and current_item:
                existing = current_item.get("description_continuation")
                current_item["description_continuation"] = (
                    f"{existing}\n{service_name}".strip() if existing else service_name
                )

    return items


def select_line_items_for_total(
    line_items: list[dict[str, Any]],
    total_value: float | None,
    tolerance: float = 0.01,
) -> list[dict[str, Any]]:
    if not line_items:
        return []

    sections: dict[int, list[dict[str, Any]]] = {}
    for item in line_items:
        section = int(item.get("quote_section_index") or 1)
        sections.setdefault(section, []).append(item)

    if len(sections) <= 1 or total_value is None:
        return _reindex_items(line_items)

    target_total = float(total_value)
    matched_sections: list[tuple[float, int]] = []
    for section, items in sections.items():
        section_total = sum(float(item.get("net_total_value") or 0.0) for item in items)
        diff = abs(target_total - section_total)
        if diff <= tolerance:
            matched_sections.append((diff, section))

    if not matched_sections:
        return _reindex_items(line_items)

    matched_sections.sort(key=lambda row: (row[0], row[1]))
    selected_section = matched_sections[0][1]
    return _reindex_items(sections[selected_section])


def _reindex_items(line_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for idx, item in enumerate(line_items, start=1):
        copied = dict(item)
        copied["item_index"] = idx
        normalized.append(copied)
    return normalized


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
