from pdf_quote_extractor.normalize import (
    parse_currency_code,
    parse_currency_value,
    parse_date_value,
    parse_number_value,
    split_term_range,
)


def test_parse_currency_value_and_code():
    assert parse_currency_value("USD 617,572.80") == 617572.80
    assert parse_currency_value("$205,857.60") == 205857.60
    assert parse_currency_value("EUR 617.572,80") == 617572.80
    assert parse_currency_value("USD 205857,60") == 205857.60
    assert parse_currency_code("USD 617,572.80") == "USD"
    assert parse_currency_code("EUR 99.10") == "EUR"
    assert parse_currency_code("\u20ac 99,10") == "EUR"


def test_parse_number_value_accepts_comma_or_dot_decimal():
    assert parse_number_value("1.17") == 1.17
    assert parse_number_value("1,17") == 1.17
    assert parse_number_value("10,000", allow_thousands=True) == 10000
    assert parse_number_value("10.000,50", allow_thousands=True) == 10000.50


def test_parse_date_value():
    assert parse_date_value("12/30/2025", ["%m/%d/%Y"]) == "2025-12-30"


def test_split_term_range():
    start, end = split_term_range("12/31/2025 - 12/30/2028")
    assert start == "12/31/2025"
    assert end == "12/30/2028"
