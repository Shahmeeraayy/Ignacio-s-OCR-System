from pdf_quote_extractor.business import parse_line_items


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

