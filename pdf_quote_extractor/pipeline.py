from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .business import extract_business_summary, parse_line_items, select_line_items_for_total
from .config import load_config
from .io_utils import gather_pdfs
from .raw_extract import extract_pdf_raw
from .template_fill import (
    DEFAULT_TEMPLATE_SHEET,
    fill_quote_template,
)
from .validate import run_validation
from .writers import write_excel, write_json


def _resolve_input_pdfs(input_paths: Path | list[Path]) -> list[Path]:
    if isinstance(input_paths, Path):
        candidates = [input_paths]
    else:
        candidates = list(input_paths)
    if not candidates:
        raise ValueError("No input paths were provided.")

    all_pdfs: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        for pdf in gather_pdfs(candidate):
            normalized = pdf.resolve()
            if normalized in seen:
                continue
            seen.add(normalized)
            all_pdfs.append(pdf)
    return sorted(all_pdfs)


def _error_rows(file_name: str, path: Path, message: str) -> dict[str, list[dict[str, Any]]]:
    metadata = {
        "file": file_name,
        "path": str(path),
        "pages": None,
        "creator": None,
        "producer": None,
        "creation_date": None,
        "is_encrypted": None,
        "parse_timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
    }
    business = {
        "file": file_name,
        "error": message,
    }
    validation = {
        "file": file_name,
        "rule_id": "processing_error",
        "severity": "critical",
        "status": "FAIL",
        "observed_value": message,
        "expected_value": None,
        "details": "Unhandled exception while processing PDF.",
    }
    return {
        "document_metadata": [metadata],
        "business_summary": [business],
        "validation_report": [validation],
    }


def process_one_pdf(
    pdf_path: Path,
    config: dict[str, Any],
    ocr_mode: str,
    include_char_layer: bool,
    include_tables: bool,
    strict: bool,
    tesseract_cmd: str | None,
    poppler_path: str | None,
) -> tuple[dict[str, Any], bool]:
    file_name = pdf_path.name
    try:
        raw_result = extract_pdf_raw(
            pdf_path=pdf_path,
            config=config,
            ocr_mode=ocr_mode,
            include_char_layer=include_char_layer,
            include_tables=include_tables,
            tesseract_cmd=tesseract_cmd,
            poppler_path=poppler_path,
        )
        line_items = parse_line_items(
            file_name=file_name,
            tables_structured=raw_result["tables_structured"],
            line_item_rules=config.get("line_item_rules", {}),
        )
        business_summary = extract_business_summary(
            file_name=file_name,
            full_text=raw_result.get("full_text", ""),
            tables_structured=raw_result.get("tables_structured", []),
            line_items=line_items,
            config=config,
        )
        reconciled_line_items = select_line_items_for_total(
            line_items=line_items,
            total_value=business_summary.get("total_value"),
            tolerance=float(config.get("validation", {}).get("money_tolerance", 0.01)),
        )
        business_summary["line_items_total_value"] = (
            sum(item.get("net_total_value") or 0.0 for item in reconciled_line_items)
            if reconciled_line_items
            else None
        )
        validation_rows, critical_failed = run_validation(
            file_name=file_name,
            raw_result=raw_result,
            line_items=line_items,
            business_summary=business_summary,
            config=config,
        )
    except Exception as exc:
        error_payload = _error_rows(file_name=file_name, path=pdf_path, message=str(exc))
        per_file_json = {
            "metadata": error_payload["document_metadata"][0],
            "pages": [],
            "text_lines": [],
            "text_words": [],
            "tables_raw": [],
            "line_items_parsed": [],
            "links": [],
            "images": [],
            "business_summary": error_payload["business_summary"][0],
            "validation_report": error_payload["validation_report"],
            "error": str(exc),
        }
        return per_file_json, True

    per_file_json = {
        "metadata": raw_result["metadata"],
        "pages": raw_result["pages"],
        "text_lines": raw_result["text_lines"],
        "text_words": raw_result["text_words"],
        "tables_raw": raw_result["tables_raw"],
        "line_items_parsed": line_items,
        "links": raw_result["links"],
        "images": raw_result["images"],
        "business_summary": business_summary,
        "validation_report": validation_rows,
        "ocr_pages": raw_result.get("ocr_pages", []),
    }
    if include_char_layer:
        per_file_json["text_chars"] = raw_result.get("text_chars", [])

    return per_file_json, (strict and critical_failed)


def run_pipeline(
    input_path: Path | list[Path],
    output_path: Path,
    json_output_path: Path,
    config_path: Path,
    ocr_mode: str,
    strict: bool,
    include_char_layer: bool,
    include_tables: bool,
    tesseract_cmd: str | None,
    poppler_path: str | None,
    template_path: Path | None = None,
    template_output_path: Path | None = None,
    euro_rate: float | None = None,
    margin_percent: float | None = None,
    template_sheet: str = DEFAULT_TEMPLATE_SHEET,
    template_header_row: int | None = None,
    template_data_start_row: int | None = None,
    write_audit_workbook: bool = True,
) -> tuple[int, dict[str, Any]]:
    config = load_config(config_path)
    pdfs = _resolve_input_pdfs(input_path)
    if not pdfs:
        raise ValueError(f"No PDF files found under: {input_path}")

    rows_by_sheet: dict[str, list[dict[str, Any]]] = {
        "document_metadata": [],
        "pages": [],
        "text_lines": [],
        "text_words": [],
        "text_chars": [],
        "tables_raw": [],
        "line_items_parsed": [],
        "links": [],
        "images": [],
        "business_summary": [],
        "validation_report": [],
    }
    files_payload: list[dict[str, Any]] = []
    strict_failures = False

    for pdf_path in pdfs:
        per_file_json, is_strict_failure = process_one_pdf(
            pdf_path=pdf_path,
            config=config,
            ocr_mode=ocr_mode,
            include_char_layer=include_char_layer,
            include_tables=include_tables,
            strict=strict,
            tesseract_cmd=tesseract_cmd,
            poppler_path=poppler_path,
        )
        files_payload.append(per_file_json)
        strict_failures = strict_failures or is_strict_failure

        rows_by_sheet["document_metadata"].append(per_file_json["metadata"])
        rows_by_sheet["pages"].extend(per_file_json.get("pages", []))
        rows_by_sheet["text_lines"].extend(per_file_json.get("text_lines", []))
        rows_by_sheet["text_words"].extend(per_file_json.get("text_words", []))
        rows_by_sheet["tables_raw"].extend(per_file_json.get("tables_raw", []))
        rows_by_sheet["line_items_parsed"].extend(per_file_json.get("line_items_parsed", []))
        rows_by_sheet["links"].extend(per_file_json.get("links", []))
        rows_by_sheet["images"].extend(per_file_json.get("images", []))
        rows_by_sheet["business_summary"].append(per_file_json.get("business_summary", {}))
        rows_by_sheet["validation_report"].extend(per_file_json.get("validation_report", []))
        if include_char_layer:
            rows_by_sheet["text_chars"].extend(per_file_json.get("text_chars", []))

    if write_audit_workbook:
        write_excel(output_path=output_path, rows_by_sheet=rows_by_sheet, include_char_layer=include_char_layer)
    if isinstance(input_path, Path):
        input_repr: str | list[str] = str(input_path)
    else:
        input_repr = [str(value) for value in input_path]

    json_payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
        "input": input_repr,
        "file_count": len(files_payload),
        "files": files_payload,
    }
    if template_path and template_output_path:
        template_result = fill_quote_template(
            template_path=template_path,
            template_output_path=template_output_path,
            files_payload=files_payload,
            euro_rate=euro_rate,
            margin_percent=margin_percent,
            sheet_name=template_sheet,
            header_row=template_header_row,
            data_start_row=template_data_start_row,
        )
        json_payload["template_output"] = template_result
    write_json(json_output_path, json_payload)

    exit_code = 2 if strict_failures else 0
    return exit_code, json_payload
