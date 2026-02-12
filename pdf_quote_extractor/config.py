from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml


DEFAULT_CONFIG: dict[str, Any] = {
    "field_patterns": {
        "quote_number": [r"Quote\s*#:\s*([A-Z0-9\-]+)"],
        "expiration_date": [r"Expiration Date:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})"],
        "subscription_period": [r"Subscription Period:\s*([^\n\r]+)"],
        "payment_method": [r"Payment Method:\s*([^\n\r]+)"],
        "total_raw": [r"\bTOTAL:\s*([A-Z$][A-Z$\s0-9,.\-]+)"],
        "overall_total_raw": [r"Overall Total:\s*([$A-Z\s0-9,.\-]+)"],
        "payment_year_1_raw": [r"Payment Year 1:\s*([$0-9,.\-]+)"],
        "payment_year_2_raw": [r"Payment Year 2:\s*([$0-9,.\-]+)"],
        "payment_year_3_raw": [r"Payment Year 3:\s*([$0-9,.\-]+)"],
        "regional_director": [r"REGIONAL DIRECTOR[\s\S]{0,250}\n([^\n]+)"],
        "regional_director_email": [r"([a-zA-Z0-9._%+\-]+@netskope\.com)"],
        "payment_terms": [r"\b(Net\s*[0-9]+)\b"],
    },
    "table_settings": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
    },
    "line_item_rules": {
        "header_contains": [
            "Service/Product Name",
            "Service/Product",
            "Code/SKU",
            "Subscription",
            "Term",
            "List Unit Price",
            "Discount",
            "Net Unit Price",
            "Net Total",
        ],
        "min_columns": 8,
    },
    "normalization": {
        "currency_default": "USD",
        "date_input_formats": ["%m/%d/%Y", "%m/%d/%y"],
    },
    "validation": {
        "money_tolerance": 0.01,
        "critical_rules": [
            "table_presence",
            "line_item_presence",
            "total_vs_line_items",
            "total_vs_overall_total",
        ],
    },
    "ocr": {
        "min_native_text_chars": 50,
        "dpi": 300,
    },
}


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged: dict[str, Any] = dict(base)
    for key, value in override.items():
        if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
            merged[key] = _deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged


def load_config(config_path: Path) -> dict[str, Any]:
    data: dict[str, Any] = {}
    if config_path.exists():
        with config_path.open("r", encoding="utf-8") as handle:
            parsed = yaml.safe_load(handle) or {}
            if isinstance(parsed, dict):
                data = parsed
    return _deep_merge(DEFAULT_CONFIG, data)

