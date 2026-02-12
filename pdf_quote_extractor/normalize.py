from __future__ import annotations

import re
from datetime import datetime
from typing import Iterable


MONEY_PATTERN = re.compile(r"([$€£]|USD|EUR|GBP)?\s*([-+]?[0-9][0-9,]*(?:\.[0-9]{1,4})?)")


def parse_currency_value(raw: str | None) -> float | None:
    if not raw:
        return None
    match = MONEY_PATTERN.search(raw.replace("\n", " "))
    if not match:
        return None
    amount = match.group(2).replace(",", "")
    try:
        return float(amount)
    except ValueError:
        return None


def parse_currency_code(raw: str | None, default_currency: str = "USD") -> str | None:
    if not raw:
        return None
    upper = raw.upper()
    if "USD" in upper or "$" in raw:
        return "USD"
    if "EUR" in upper or "€" in raw:
        return "EUR"
    if "GBP" in upper or "£" in raw:
        return "GBP"
    return default_currency


def parse_date_value(raw: str | None, date_formats: Iterable[str]) -> str | None:
    if not raw:
        return None
    for fmt in date_formats:
        try:
            return datetime.strptime(raw.strip(), fmt).date().isoformat()
        except ValueError:
            continue
    return None


def split_term_range(raw: str | None) -> tuple[str | None, str | None]:
    if not raw:
        return None, None
    match = re.search(
        r"([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})\s*-\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        raw,
    )
    if not match:
        return None, None
    return match.group(1), match.group(2)

