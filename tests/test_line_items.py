from pdf_quote_extractor.business import parse_line_items, select_line_items_for_total


def test_parse_line_items_with_continuation():
    tables_structured = [
        {
            "file": "sample.pdf",
            "page": 1,
            "table_index": 1,
            "rows": [
                [
                    "Service/Product Name",
                    "Service/Product\nCode/SKU",
                    "Subscription\nUnits/\nQuantity",
                    "Term",
                    "List Unit Price",
                    "Discount\n(%)",
                    "Net Unit Price",
                    "Net Total",
                ],
                [
                    "Secure Web Gateway\nStandard",
                    "NK-P-SWG-STD",
                    "10,000\nUsers",
                    "12/31/2025 -\n12/30/2028",
                    "USD 166.17",
                    "76.51",
                    "USD 39.03",
                    "USD 390,300.00",
                ],
                [
                    "This is a continuation description row.",
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                ],
            ],
        }
    ]
    rules = {"header_contains": ["Service/Product Name", "Code/SKU"], "min_columns": 8}
    items = parse_line_items("sample.pdf", tables_structured, rules)
    assert len(items) == 1
    assert items[0]["sku"] == "NK-P-SWG-STD"
    assert items[0]["net_total_value"] == 390300.00
    assert "continuation description" in (items[0]["description_continuation"] or "").lower()


def test_select_line_items_for_total_uses_matching_section():
    line_items = [
        {"item_index": 1, "quote_section_index": 1, "net_total_value": 100.0, "sku": "A"},
        {"item_index": 2, "quote_section_index": 1, "net_total_value": 50.0, "sku": "B"},
        {"item_index": 3, "quote_section_index": 2, "net_total_value": 20.0, "sku": "C"},
    ]

    selected = select_line_items_for_total(line_items=line_items, total_value=150.0, tolerance=0.01)

    assert len(selected) == 2
    assert selected[0]["sku"] == "A"
    assert selected[1]["sku"] == "B"
    assert selected[0]["item_index"] == 1
    assert selected[1]["item_index"] == 2


def test_parse_line_items_with_reordered_columns():
    tables_structured = [
        {
            "file": "sample.pdf",
            "page": 1,
            "table_index": 1,
            "rows": [
                [
                    "Service/Product\nCode/SKU",
                    "Service/Product Name",
                    "Net Total",
                    "Subscription\nUnits/\nQuantity",
                    "Term",
                    "List Unit Price",
                    "Discount\n(%)",
                    "Net Unit Price",
                ],
                [
                    "NK-TEST-001",
                    "Custom Product",
                    "USD 900.00",
                    "10 Users",
                    "12/01/2025 - 12/01/2026",
                    "USD 100.00",
                    "10.00",
                    "USD 90.00",
                ],
                ["Extra description line", None, None, None, None, None, None, None],
            ],
        }
    ]
    rules = {"header_contains": ["Service/Product Name", "Code/SKU"], "min_columns": 8}
    items = parse_line_items("sample.pdf", tables_structured, rules)

    assert len(items) == 1
    assert items[0]["service_name"] == "Custom Product"
    assert items[0]["sku"] == "NK-TEST-001"
    assert items[0]["units_qty"] == "10 Users"
    assert items[0]["term_start"] == "12/01/2025"
    assert items[0]["term_end"] == "12/01/2026"
    assert items[0]["list_unit_price_value"] == 100.0
    assert items[0]["discount_pct_value"] == 10.0
    assert items[0]["net_unit_price_value"] == 90.0
    assert items[0]["net_total_value"] == 900.0
    assert "extra description line" in (items[0]["description_continuation"] or "").lower()
