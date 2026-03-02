from __future__ import annotations

import re
from datetime import datetime
from typing import Iterable


MONEY_PATTERN = re.compile(r"([$\u20ac\u00a3]|USD|EUR|GBP)?\s*([-+]?[0-9][0-9.,]*)", re.IGNORECASE)
NUMBER_TOKEN_PATTERN = re.compile(r"[-+]?[0-9][0-9.,]*")


def _normalize_single_separator(value: str, separator: str, *, allow_thousands: bool) -> str:
    parts = value.split(separator)
    if len(parts) == 1:
        return value

    if len(parts) > 2:
        if allow_thousands and all(part.isdigit() and len(part) == 3 for part in parts[1:]):
            return "".join(parts)
        head = "".join(parts[:-1])
        tail = parts[-1]
        return f"{head}.{tail}" if tail else head

    head, tail = parts
    if not head or not tail:
        return f"{head}{tail}"

    is_thousands_style = (
        allow_thousands
        and len(tail) == 3
        and len(head) <= 3
        and not head.startswith("0")
    )
    if is_thousands_style:
        return f"{head}{tail}"
    return f"{head}.{tail}"


def _normalize_numeric_token(token: str, *, allow_thousands: bool) -> str | None:
    cleaned = token.strip().replace(" ", "").replace("\u00a0", "").replace("'", "")
    if not cleaned:
        return None

    sign = ""
    if cleaned[0] in "+-":
        sign = cleaned[0]
        cleaned = cleaned[1:]

    if not cleaned or not any(char.isdigit() for char in cleaned):
        return None

    comma_count = cleaned.count(",")
    dot_count = cleaned.count(".")

    if comma_count and dot_count:
        decimal_separator = "," if cleaned.rfind(",") > cleaned.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        cleaned = cleaned.replace(thousands_separator, "")
        if decimal_separator == ",":
            cleaned = cleaned.replace(",", ".")
    elif comma_count:
        cleaned = _normalize_single_separator(cleaned, ",", allow_thousands=allow_thousands)
    elif dot_count:
        cleaned = _normalize_single_separator(cleaned, ".", allow_thousands=allow_thousands)

    normalized = f"{sign}{cleaned}"
    if normalized in {"", "+", "-", ".", "+.", "-."}:
        return None
    return normalized


def parse_number_value(raw: str | None, *, allow_thousands: bool = True) -> float | None:
    if not raw:
        return None
    match = NUMBER_TOKEN_PATTERN.search(str(raw).replace("\n", " "))
    if not match:
        return None
    normalized = _normalize_numeric_token(match.group(0), allow_thousands=allow_thousands)
    if normalized is None:
        return None
    try:
        return float(normalized)
    except ValueError:
        return None


def parse_currency_value(raw: str | None) -> float | None:
    if not raw:
        return None
    text = str(raw).replace("\n", " ")
    match = MONEY_PATTERN.search(text)
    candidate = match.group(2) if match else text
    return parse_number_value(candidate, allow_thousands=True)


def parse_currency_code(raw: str | None, default_currency: str = "USD") -> str | None:
    if not raw:
        return None
    upper = raw.upper()
    if "USD" in upper or "$" in raw:
        return "USD"
    if "EUR" in upper or "\u20ac" in raw or "â‚¬" in raw:
        return "EUR"
    if "GBP" in upper or "\u00a3" in raw or "Â£" in raw:
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
