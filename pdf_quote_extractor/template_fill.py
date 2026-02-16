from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .normalize import parse_currency_value


DEFAULT_TEMPLATE_SHEET = "QuoteExportResults"
DEFAULT_TEMPLATE_HEADER_ROW = 4
DEFAULT_TEMPLATE_DATA_START_ROW = 5

TARGET_HEADERS = [
    "Date",
    "Expires",
    "ExpectedClose",
    "Item",
    "Quantity",
    "Salesprice",
    "Salesdiscount",
    "Purchaseprice",
    "PurchaseDiscount",
    "ContractStart",
    "ContractEnd",
    "Serial#Supported",
    "Rebate",
    "Opportunity",
    "Memo (Line)",
    "Quote ID (Line)",
]

HEADER_NUMBER_FORMATS: dict[str, str] = {
    "Quantity": "0",
    "Salesprice": "0.00",
    "Salesdiscount": "0.00%",
    "Purchaseprice": "0.00",
    "PurchaseDiscount": "0.00%",
}

HEADER_ALIASES: dict[str, set[str]] = {
    "Date": {"date"},
    "Expires": {"expires"},
    "ExpectedClose": {"expectedclose"},
    "Item": {"item"},
    "Quantity": {"quantity", "qty"},
    "Salesprice": {"salesprice", "sales_price", "sales price"},
    "Salesdiscount": {"salesdiscount", "sales_discount", "sales discount"},
    "Purchaseprice": {"purchaseprice", "purchase_price", "purchase price"},
    "PurchaseDiscount": {"purchasediscount", "purchase_discount", "purchase discount"},
    "ContractStart": {"contractstart", "contract_start", "contract start"},
    "ContractEnd": {"contractend", "contract_end", "contract end"},
    "Serial#Supported": {"serialsupported", "serial", "serialnumbersupported"},
    "Rebate": {"rebate"},
    "Opportunity": {"opportunity"},
    "Memo (Line)": {"memoline", "memo"},
    "Quote ID (Line)": {"quoteidline", "quoteid", "quote id"},
}

NORMALIZED_HEADER_ALIASES: dict[str, set[str]] = {
    target: {re.sub(r"[^a-z0-9]+", "", alias.strip().lower()) for alias in aliases}
    for target, aliases in HEADER_ALIASES.items()
}


def _parse_creation_date(raw: str | None) -> str | None:
    if not raw:
        return None
    match = re.search(r"D:(\d{4})(\d{2})(\d{2})", raw)
    if not match:
        return None
    try:
        year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))
        return datetime(year, month, day).strftime("%m/%d/%Y")
    except ValueError:
        return None


def _clean_single_line(text: str | None) -> str | None:
    if text is None:
        return None
    cleaned = " ".join(str(text).replace("\n", " ").split())
    return cleaned or None


def _normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text.strip().lower())


def _parse_quantity(raw: str | None) -> float | int | None:
    if not raw:
        return None
    match = re.search(r"([-+]?[0-9][0-9,]*(?:\.[0-9]+)?)", raw.replace("\n", " "))
    if not match:
        return None
    numeric = match.group(1).replace(",", "")
    try:
        value = float(numeric)
    except ValueError:
        return None
    if value.is_integer():
        return int(value)
    return value


def _parse_discount_fraction(
    discount_pct_value: float | None,
    discount_pct_raw: str | None,
) -> float | None:
    if discount_pct_value is not None:
        return round(float(discount_pct_value) / 100.0, 6)
    if discount_pct_raw:
        match = re.search(r"[-+]?[0-9]+(?:\.[0-9]+)?", discount_pct_raw)
        if match:
            try:
                return round(float(match.group(0)) / 100.0, 6)
            except ValueError:
                return None
    return None


def _parse_sales_discount(
    sales_price: float | None,
    purchase_price: float | None,
    euro_rate: float,
    margin_percent: float,
) -> float | None:
    if sales_price is None or purchase_price is None:
        return None
    if sales_price == 0:
        return None
    if euro_rate <= 0:
        return None
    margin_multiplier = 1.0 + (margin_percent / 100.0)
    value = 1.0 - (((purchase_price / euro_rate) * margin_multiplier) / sales_price)
    return round(value, 6)


def _parse_net_total(item: dict[str, Any]) -> float | None:
    value = item.get("net_total_value")
    if value is not None:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
    return parse_currency_value(item.get("net_total_raw"))


def _should_skip_template_item(
    item: dict[str, Any],
    sales_price: float | None,
    net_total: float | None,
) -> bool:
    # Exclude non-billable/included lines from template output.
    included_text = str(item.get("net_unit_price_raw") or "").strip().lower()
    if included_text == "included":
        return True
    if net_total is not None and abs(net_total) <= 1e-9:
        if sales_price is None or abs(sales_price) <= 1e-9:
            return True
    return False


def _build_template_rows(
    files_payload: list[dict[str, Any]],
    euro_rate: float,
    margin_percent: float,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for file_payload in files_payload:
        metadata = file_payload.get("metadata", {})
        summary = file_payload.get("business_summary", {})
        line_items = file_payload.get("line_items_parsed", [])

        quote_id = summary.get("quote_number")
        expiration_date = summary.get("expiration_date")
        creation_date = _parse_creation_date(metadata.get("creation_date"))

        for item in line_items:
            discount_fraction = _parse_discount_fraction(
                item.get("discount_pct_value"),
                item.get("discount_pct_raw"),
            )
            purchase_price = item.get("net_unit_price_value")
            if purchase_price is None:
                purchase_price = parse_currency_value(item.get("net_unit_price_raw"))
            sales_price = item.get("list_unit_price_value")
            if sales_price is None:
                sales_price = parse_currency_value(item.get("list_unit_price_raw"))
            net_total = _parse_net_total(item)
            if _should_skip_template_item(item=item, sales_price=sales_price, net_total=net_total):
                continue
            sales_discount = _parse_sales_discount(
                sales_price=sales_price,
                purchase_price=purchase_price,
                euro_rate=euro_rate,
                margin_percent=margin_percent,
            )

            row = {
                "Date": creation_date,
                "Expires": expiration_date,
                "ExpectedClose": expiration_date,
                "Item": _clean_single_line(item.get("sku")),
                "Quantity": _parse_quantity(item.get("units_qty")),
                "Salesprice": sales_price,
                "Salesdiscount": sales_discount,
                "Purchaseprice": purchase_price,
                "PurchaseDiscount": discount_fraction,
                "ContractStart": item.get("term_start"),
                "ContractEnd": item.get("term_end"),
                "Serial#Supported": None,
                "Rebate": None,
                "Opportunity": quote_id,
                "Memo (Line)": None,
                "Quote ID (Line)": quote_id,
            }
            rows.append(row)
    return rows


def _match_headers_in_row(ws, header_row: int) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col_idx).value
        if isinstance(value, str):
            normalized = _normalize_header(value)
            for target_header, aliases in NORMALIZED_HEADER_ALIASES.items():
                if normalized in aliases and target_header not in headers:
                    headers[target_header] = col_idx
                    break
    return headers


def _resolve_header_row_and_columns(ws, header_row: int | None) -> tuple[int, dict[str, int]]:
    if header_row is not None:
        headers = _match_headers_in_row(ws, header_row)
        if all(name in headers for name in TARGET_HEADERS):
            return header_row, headers

    max_scan_row = min(ws.max_row, 50)
    best_row = None
    best_headers: dict[str, int] = {}
    for row_idx in range(1, max_scan_row + 1):
        headers = _match_headers_in_row(ws, row_idx)
        if len(headers) > len(best_headers):
            best_headers = headers
            best_row = row_idx
        if all(name in headers for name in TARGET_HEADERS):
            return row_idx, headers

    missing = [name for name in TARGET_HEADERS if name not in best_headers]
    missing_str = ", ".join(missing)
    if best_row is not None:
        raise ValueError(
            f"Template headers not fully found. Best match row={best_row}, missing: {missing_str}"
        )
    raise ValueError(f"Template headers not found. Missing: {missing_str}")


def fill_quote_template(
    template_path: Path,
    template_output_path: Path,
    files_payload: list[dict[str, Any]],
    euro_rate: float | None,
    margin_percent: float | None,
    sheet_name: str = DEFAULT_TEMPLATE_SHEET,
    header_row: int | None = None,
    data_start_row: int | None = None,
) -> dict[str, Any]:
    if euro_rate is None or euro_rate <= 0:
        raise ValueError("euro_rate must be provided and greater than 0.")
    if margin_percent is None:
        raise ValueError("margin_percent must be provided.")

    wb = load_workbook(template_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Template sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    resolved_header_row, header_columns = _resolve_header_row_and_columns(ws, header_row)

    if data_start_row is None:
        resolved_data_start_row = resolved_header_row + 1
    else:
        resolved_data_start_row = data_start_row
    if resolved_data_start_row <= resolved_header_row:
        raise ValueError(
            f"data_start_row ({resolved_data_start_row}) must be greater than header_row ({resolved_header_row})."
        )

    rows_to_write = _build_template_rows(
        files_payload=files_payload,
        euro_rate=float(euro_rate),
        margin_percent=float(margin_percent),
    )
    capacity = ws.max_row - resolved_data_start_row + 1
    if len(rows_to_write) > capacity:
        raise ValueError(
            f"Template row capacity exceeded: {len(rows_to_write)} rows required, capacity is {capacity}."
        )

    # Clear previous data only in yellow target columns.
    for row_idx in range(resolved_data_start_row, ws.max_row + 1):
        for col_idx in header_columns.values():
            ws.cell(row_idx, col_idx).value = None

    # Fill rows from extracted line items.
    for offset, row_values in enumerate(rows_to_write):
        row_idx = resolved_data_start_row + offset
        for header, col_idx in header_columns.items():
            cell = ws.cell(row_idx, col_idx)
            value = row_values.get(header)
            cell.value = value
            number_format = HEADER_NUMBER_FORMATS.get(header)
            if number_format and value not in (None, ""):
                cell.number_format = number_format

    wb.save(template_output_path)
    return {
        "template_path": str(template_path),
        "template_output_path": str(template_output_path),
        "sheet_name": sheet_name,
        "rows_written": len(rows_to_write),
        "capacity": capacity,
        "euro_rate": float(euro_rate),
        "margin_percent": float(margin_percent),
        "header_row": resolved_header_row,
        "data_start_row": resolved_data_start_row,
    }
