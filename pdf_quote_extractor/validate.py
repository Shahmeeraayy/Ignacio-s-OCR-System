from __future__ import annotations

from typing import Any


def _status(ok: bool) -> str:
    return "PASS" if ok else "FAIL"


def run_validation(
    file_name: str,
    raw_result: dict[str, Any],
    line_items: list[dict[str, Any]],
    business_summary: dict[str, Any],
    config: dict[str, Any],
) -> tuple[list[dict[str, Any]], bool]:
    validation_cfg = config.get("validation", {})
    tolerance = float(validation_cfg.get("money_tolerance", 0.01))
    critical_rules = set(validation_cfg.get("critical_rules", []))

    report: list[dict[str, Any]] = []

    def add(
        rule_id: str,
        severity: str,
        ok: bool,
        observed: Any,
        expected: Any,
        details: str,
    ) -> None:
        report.append(
            {
                "file": file_name,
                "rule_id": rule_id,
                "severity": severity,
                "status": _status(ok),
                "observed_value": observed,
                "expected_value": expected,
                "details": details,
            }
        )

    pages = raw_result.get("pages", [])
    add(
        "page_presence",
        "critical",
        bool(pages),
        len(pages),
        ">0",
        "At least one page must be extracted.",
    )

    all_words = raw_result.get("text_words", [])
    add(
        "word_presence",
        "high",
        bool(all_words),
        len(all_words),
        ">0",
        "Word layer should not be empty.",
    )

    per_page_word_zeros = [page["page"] for page in pages if int(page.get("word_count", 0)) == 0]
    add(
        "word_coverage_per_page",
        "high",
        not per_page_word_zeros,
        ",".join(str(page) for page in per_page_word_zeros) if per_page_word_zeros else None,
        "none",
        "Each page should have at least one extracted word.",
    )

    tables = raw_result.get("tables_raw", [])
    add(
        "table_presence",
        "critical",
        bool(tables),
        len(tables),
        ">0",
        "At least one table cell expected for quote-style PDFs.",
    )

    add(
        "line_item_presence",
        "critical",
        bool(line_items),
        len(line_items),
        ">0",
        "At least one parsed line item expected.",
    )

    total_value = business_summary.get("total_value")
    line_total = business_summary.get("line_items_total_value")
    if total_value is not None and line_total is not None:
        diff = abs(float(total_value) - float(line_total))
        add(
            "total_vs_line_items",
            "critical",
            diff <= tolerance,
            round(diff, 6),
            f"<= {tolerance}",
            "TOTAL should reconcile with sum of parsed line item net totals.",
        )
    else:
        add(
            "total_vs_line_items",
            "critical",
            False,
            f"total={total_value}, line_total={line_total}",
            "both numeric",
            "Unable to reconcile totals because at least one value is missing.",
        )

    overall_total = business_summary.get("overall_total_value")
    if total_value is not None and overall_total is not None:
        diff = abs(float(total_value) - float(overall_total))
        add(
            "total_vs_overall_total",
            "critical",
            diff <= tolerance,
            round(diff, 6),
            f"<= {tolerance}",
            "TOTAL and Overall Total should match.",
        )
    else:
        add(
            "total_vs_overall_total",
            "critical",
            False,
            f"total={total_value}, overall_total={overall_total}",
            "both numeric",
            "Unable to compare TOTAL and Overall Total.",
        )

    links = raw_result.get("links", [])
    add(
        "link_capture",
        "medium",
        True,
        len(links),
        "n/a",
        "Link extraction is informational and non-blocking.",
    )

    critical_failed = any(
        row["rule_id"] in critical_rules and row["status"] == "FAIL"
        for row in report
    )
    return report, critical_failed

