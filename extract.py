from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pdf_quote_extractor import run_pipeline
from pdf_quote_extractor.normalize import parse_number_value


def _default_json_output(output_path: Path) -> Path:
    if output_path.suffix:
        return output_path.with_suffix(".json")
    return Path(f"{output_path}.json")


def _default_template_output(output_path: Path) -> Path:
    stem = output_path.stem if output_path.suffix else str(output_path)
    return output_path.with_name(f"{stem}.template_filled.xlsx")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lossless PDF extraction for Netskope quote PDFs to Excel and JSON."
    )
    parser.add_argument(
        "--input",
        required=True,
        nargs="+",
        help="One or more PDF files and/or folders.",
    )
    parser.add_argument("--output", default="output.xlsx", help="Output Excel workbook path.")
    parser.add_argument(
        "--json-output",
        default=None,
        help="Output JSON path. Defaults to same name as --output with .json extension.",
    )
    parser.add_argument("--config", default="config.yaml", help="YAML config path.")
    parser.add_argument(
        "--ocr-mode",
        choices=["auto", "off", "always"],
        default="auto",
        help="OCR behavior: auto fallback, disabled, or always OCR.",
    )
    parser.add_argument(
        "--strict",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Enable strict validation (default: true).",
    )
    parser.add_argument(
        "--char-layer",
        action="store_true",
        help="Enable optional character-level layer.",
    )
    parser.add_argument(
        "--no-char-layer",
        action="store_true",
        help="Disable character-level layer (default behavior).",
    )
    parser.add_argument(
        "--no-tables",
        action="store_true",
        help="Disable table extraction.",
    )
    parser.add_argument(
        "--tesseract-cmd",
        default=None,
        help="Path to tesseract executable if not in PATH.",
    )
    parser.add_argument(
        "--poppler-path",
        default=None,
        help="Optional Poppler bin directory path for OCR on Windows.",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Optional Excel template path. If set, yellow fields are populated from extracted data.",
    )
    parser.add_argument(
        "--template-output",
        default=None,
        help="Output path for filled template workbook. Default: <output>.template_filled.xlsx",
    )
    parser.add_argument(
        "--template-sheet",
        default="QuoteExportResults",
        help="Template sheet name to populate (default: QuoteExportResults).",
    )
    parser.add_argument(
        "--template-header-row",
        type=int,
        default=None,
        help="Header row index in the template (auto-detected if omitted).",
    )
    parser.add_argument(
        "--template-data-start-row",
        type=int,
        default=None,
        help="First data row index in the template (defaults to header_row + 1).",
    )
    parser.add_argument(
        "--template-only",
        action="store_true",
        help="Only generate the filled template (skip audit workbook with extra sheets/columns).",
    )
    parser.add_argument(
        "--euro-rate",
        type=str,
        default=None,
        help="Euro exchange rate used in Salesdiscount formula (required when template output is enabled).",
    )
    parser.add_argument(
        "--margin-percent",
        type=str,
        default=None,
        help="Margin percent used in Salesdiscount formula (e.g. 10 for 10%%). Required when template output is enabled.",
    )

    args = parser.parse_args()

    input_paths = [Path(value) for value in args.input]
    output_path = Path(args.output)
    json_output_path = Path(args.json_output) if args.json_output else _default_json_output(output_path)
    config_path = Path(args.config)
    template_path = Path(args.template) if args.template else None
    template_output_path = (
        Path(args.template_output)
        if args.template_output
        else (_default_template_output(output_path) if template_path else None)
    )

    include_char_layer = bool(args.char_layer and not args.no_char_layer)
    include_tables = not args.no_tables
    write_audit_workbook = not args.template_only
    parsed_euro_rate = (
        parse_number_value(args.euro_rate, allow_thousands=True) if args.euro_rate is not None else None
    )
    parsed_margin_percent = (
        parse_number_value(args.margin_percent, allow_thousands=True)
        if args.margin_percent is not None
        else None
    )
    if args.euro_rate is not None and parsed_euro_rate is None:
        print(
            "Failed: --euro-rate must be numeric and may use . or , as decimal separator.",
            file=sys.stderr,
        )
        return 1
    if args.margin_percent is not None and parsed_margin_percent is None:
        print(
            "Failed: --margin-percent must be numeric and may use . or , as decimal separator.",
            file=sys.stderr,
        )
        return 1
    if template_path and template_output_path:
        if parsed_euro_rate is None or parsed_euro_rate <= 0:
            print(
                "Failed: --euro-rate is required and must be > 0 when template output is enabled.",
                file=sys.stderr,
            )
            return 1
        if parsed_margin_percent is None:
            print(
                "Failed: --margin-percent is required when template output is enabled.",
                file=sys.stderr,
            )
            return 1

    try:
        exit_code, payload = run_pipeline(
            input_path=input_paths,
            output_path=output_path,
            json_output_path=json_output_path,
            config_path=config_path,
            ocr_mode=args.ocr_mode,
            strict=bool(args.strict),
            include_char_layer=include_char_layer,
            include_tables=include_tables,
            tesseract_cmd=args.tesseract_cmd,
            poppler_path=args.poppler_path,
            template_path=template_path,
            template_output_path=template_output_path,
            euro_rate=parsed_euro_rate,
            margin_percent=parsed_margin_percent,
            template_sheet=args.template_sheet,
            template_header_row=args.template_header_row,
            template_data_start_row=args.template_data_start_row,
            write_audit_workbook=write_audit_workbook,
        )
    except PermissionError as exc:
        print(
            f"Failed: {exc}. If the output/template file is open in Excel, close it and run again.",
            file=sys.stderr,
        )
        return 1
    except Exception as exc:
        print(f"Failed: {exc}", file=sys.stderr)
        return 1

    if write_audit_workbook:
        print(f"Wrote Excel: {output_path}")
    print(f"Wrote JSON:  {json_output_path}")
    if template_output_path:
        print(f"Wrote Template Output: {template_output_path}")
    print(f"Processed files: {payload.get('file_count', 0)}")
    if exit_code != 0:
        print("Strict validation failed for at least one file.", file=sys.stderr)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
