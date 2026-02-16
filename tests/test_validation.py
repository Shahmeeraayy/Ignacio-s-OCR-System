from pdf_quote_extractor.validate import run_validation


def test_validation_reconciliation_pass():
    raw_result = {
        "pages": [{"page": 1, "word_count": 5}],
        "text_words": [{"text": "hello"}],
        "tables_raw": [{"cell_text": "x"}],
        "links": [],
    }
    line_items = [{"net_total_value": 100.0}]
    business_summary = {
        "total_value": 100.0,
        "overall_total_value": 100.0,
        "line_items_total_value": 100.0,
    }
    config = {
        "validation": {
            "money_tolerance": 0.01,
            "critical_rules": [
                "table_presence",
                "line_item_presence",
                "total_vs_line_items",
                "total_vs_overall_total",
            ],
        }
    }
    report, critical_failed = run_validation(
        file_name="sample.pdf",
        raw_result=raw_result,
        line_items=line_items,
        business_summary=business_summary,
        config=config,
    )
    assert report
    assert critical_failed is False


def test_validation_reconciliation_fail():
    raw_result = {
        "pages": [{"page": 1, "word_count": 5}],
        "text_words": [{"text": "hello"}],
        "tables_raw": [{"cell_text": "x"}],
        "links": [],
    }
    line_items = [{"net_total_value": 90.0}]
    business_summary = {
        "total_value": 100.0,
        "overall_total_value": 100.0,
        "line_items_total_value": 90.0,
    }
    config = {
        "validation": {
            "money_tolerance": 0.01,
            "critical_rules": ["total_vs_line_items"],
        }
    }
    _, critical_failed = run_validation(
        file_name="sample.pdf",
        raw_result=raw_result,
        line_items=line_items,
        business_summary=business_summary,
        config=config,
    )
    assert critical_failed is True


def test_validation_overall_total_missing_is_non_blocking():
    raw_result = {
        "pages": [{"page": 1, "word_count": 5}],
        "text_words": [{"text": "hello"}],
        "tables_raw": [{"cell_text": "x"}],
        "links": [],
    }
    line_items = [{"net_total_value": 100.0}]
    business_summary = {
        "total_value": 100.0,
        "overall_total_value": None,
        "line_items_total_value": 100.0,
    }
    config = {
        "validation": {
            "money_tolerance": 0.01,
            "critical_rules": ["total_vs_overall_total"],
        }
    }
    report, critical_failed = run_validation(
        file_name="sample.pdf",
        raw_result=raw_result,
        line_items=line_items,
        business_summary=business_summary,
        config=config,
    )
    assert critical_failed is False
    row = next(item for item in report if item["rule_id"] == "total_vs_overall_total")
    assert row["status"] == "PASS"
