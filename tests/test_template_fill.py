from pathlib import Path
import shutil

from openpyxl import load_workbook

from pdf_quote_extractor.pipeline import run_pipeline
from pdf_quote_extractor.template_fill import _build_template_rows


def test_template_fill_for_quote_fixture(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    fixture_pdf = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_file = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"
    config_path = project_root / "config.yaml"

    output_xlsx = tmp_path / "quote_output.xlsx"
    output_json = tmp_path / "quote_output.json"
    filled_template = tmp_path / "quote_template_filled.xlsx"

    exit_code, payload = run_pipeline(
        input_path=fixture_pdf,
        output_path=output_xlsx,
        json_output_path=output_json,
        config_path=config_path,
        ocr_mode="auto",
        strict=True,
        include_char_layer=False,
        include_tables=True,
        tesseract_cmd=None,
        poppler_path=None,
        template_path=template_file,
        template_output_path=filled_template,
        euro_rate=1.17,
        margin_percent=10.0,
    )

    assert exit_code == 0
    assert payload["template_output"]["rows_written"] == 4
    assert filled_template.exists()

    wb = load_workbook(filled_template, data_only=False)
    ws = wb["QuoteExportResults"]

    # Row 5: first parsed line item.
    assert ws["C5"].value == "EUR"
    assert ws["D5"].value == "24/12/2025"
    assert ws["G5"].value == "30/12/2025"
    assert ws["H5"].value == "30/12/2025"
    assert ws["J5"].value == "Spain"
    assert ws["K5"].value == "NK-EGRESS-DIP"
    assert ws["L5"].value == 10000
    assert ws["M5"].value == 60
    assert ws["N5"].value == 0.844872
    assert ws["O5"].value == 9.9
    assert ws["P5"].value == 0.835
    assert ws["N5"].number_format == "0.00%"
    assert ws["P5"].number_format == "0.00%"
    assert ws["Q5"].value == "EXN Spain : ES Sales Stock"
    assert ws["R5"].value == "31/12/2025"
    assert ws["S5"].value == "30/12/2028"
    assert ws["V5"].value in (None, "")
    assert ws["W5"].value in (None, "")
    assert ws["X5"].value == "Q-220053-2"
    assert ws["Y5"].value == "EUR"

    # Row 6: second parsed line item.
    assert ws["L6"].value == 1
    assert ws["M6"].value == 48000
    assert ws["N6"].value == 0.50359
    assert ws["O6"].value == 25344
    assert ws["P6"].value == 0.472
    assert ws["N6"].number_format == "0.00%"
    assert ws["P6"].number_format == "0.00%"
    assert ws["K6"].value == "NK-SSLI"
    assert ws["V6"].value in (None, "")
    assert ws["W6"].value in (None, "")


def test_template_only_mode_skips_audit_workbook(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    fixture_pdf = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_file = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"
    config_path = project_root / "config.yaml"

    output_xlsx = tmp_path / "quote_output.xlsx"
    output_json = tmp_path / "quote_output.json"
    filled_template = tmp_path / "quote_template_filled.xlsx"

    exit_code, payload = run_pipeline(
        input_path=fixture_pdf,
        output_path=output_xlsx,
        json_output_path=output_json,
        config_path=config_path,
        ocr_mode="auto",
        strict=True,
        include_char_layer=False,
        include_tables=True,
        tesseract_cmd=None,
        poppler_path=None,
        template_path=template_file,
        template_output_path=filled_template,
        euro_rate=1.17,
        margin_percent=10.0,
        write_audit_workbook=False,
    )

    assert exit_code == 0
    assert payload["file_count"] == 1
    assert filled_template.exists()
    assert not output_xlsx.exists()


def test_template_fill_with_multiple_pdfs(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    fixture_pdf = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_file = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"
    config_path = project_root / "config.yaml"

    batch_dir = tmp_path / "batch"
    batch_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(fixture_pdf, batch_dir / "quote_1.pdf")
    shutil.copy2(fixture_pdf, batch_dir / "quote_2.pdf")

    output_xlsx = tmp_path / "quote_output.xlsx"
    output_json = tmp_path / "quote_output.json"
    filled_template = tmp_path / "quote_template_filled.xlsx"

    exit_code, payload = run_pipeline(
        input_path=batch_dir,
        output_path=output_xlsx,
        json_output_path=output_json,
        config_path=config_path,
        ocr_mode="off",
        strict=True,
        include_char_layer=False,
        include_tables=True,
        tesseract_cmd=None,
        poppler_path=None,
        template_path=template_file,
        template_output_path=filled_template,
        euro_rate=1.17,
        margin_percent=10.0,
        write_audit_workbook=False,
    )

    assert exit_code == 0
    assert payload["file_count"] == 2
    assert payload["template_output"]["rows_written"] == 8
    assert filled_template.exists()

    wb = load_workbook(filled_template, data_only=False)
    ws = wb["QuoteExportResults"]
    assert ws["K5"].value is not None
    assert ws["K12"].value is not None


def test_template_fill_with_multiple_input_paths(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    fixture_pdf = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_file = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"
    config_path = project_root / "config.yaml"

    pdf1 = tmp_path / "input_1.pdf"
    pdf2 = tmp_path / "input_2.pdf"
    shutil.copy2(fixture_pdf, pdf1)
    shutil.copy2(fixture_pdf, pdf2)

    output_xlsx = tmp_path / "quote_output.xlsx"
    output_json = tmp_path / "quote_output.json"
    filled_template = tmp_path / "quote_template_filled.xlsx"

    exit_code, payload = run_pipeline(
        input_path=[pdf1, pdf2],
        output_path=output_xlsx,
        json_output_path=output_json,
        config_path=config_path,
        ocr_mode="off",
        strict=True,
        include_char_layer=False,
        include_tables=True,
        tesseract_cmd=None,
        poppler_path=None,
        template_path=template_file,
        template_output_path=filled_template,
        euro_rate=1.17,
        margin_percent=10.0,
        write_audit_workbook=False,
    )

    assert exit_code == 0
    assert payload["file_count"] == 2
    assert payload["template_output"]["rows_written"] == 8


def test_template_autodetects_header_row_and_data_start(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    fixture_pdf = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    source_template = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"
    config_path = project_root / "config.yaml"

    template_file = tmp_path / "template_headers_row2.xlsx"
    shutil.copy2(source_template, template_file)

    wb = load_workbook(template_file, data_only=False)
    ws = wb["QuoteExportResults"]
    for col in range(1, 27):
        ws.cell(2, col).value = ws.cell(4, col).value
        ws.cell(4, col).value = None
    wb.save(template_file)

    output_xlsx = tmp_path / "quote_output.xlsx"
    output_json = tmp_path / "quote_output.json"
    filled_template = tmp_path / "quote_template_filled.xlsx"

    exit_code, payload = run_pipeline(
        input_path=fixture_pdf,
        output_path=output_xlsx,
        json_output_path=output_json,
        config_path=config_path,
        ocr_mode="off",
        strict=True,
        include_char_layer=False,
        include_tables=True,
        tesseract_cmd=None,
        poppler_path=None,
        template_path=template_file,
        template_output_path=filled_template,
        euro_rate=1.17,
        margin_percent=10.0,
        write_audit_workbook=False,
    )

    assert exit_code == 0
    assert payload["template_output"]["header_row"] == 2
    assert payload["template_output"]["data_start_row"] == 3

    wb_out = load_workbook(filled_template, data_only=False)
    ws_out = wb_out["QuoteExportResults"]
    assert ws_out["K3"].value is not None
    assert ws_out["M3"].value == 60
    assert ws_out["N3"].value == 0.844872
    assert ws_out["N3"].number_format == "0.00%"
    assert ws_out["P3"].number_format == "0.00%"
    assert ws_out["D3"].value == "24/12/2025"
    assert ws_out["V3"].value in (None, "")


def test_template_rows_include_included_zero_value_items():
    files_payload = [
        {
            "metadata": {"creation_date": "D:20260101000000Z"},
            "business_summary": {"quote_number": "Q-TEST", "expiration_date": "01/31/2026"},
            "line_items_parsed": [
                {
                    "sku": "SKU-A",
                    "units_qty": "1",
                    "term_start": "01/01/2026",
                    "term_end": "01/31/2026",
                    "list_unit_price_value": 100.0,
                    "discount_pct_value": 0.0,
                    "net_unit_price_value": 100.0,
                    "net_total_value": 100.0,
                },
                {
                    "sku": "SKU-B",
                    "units_qty": "1",
                    "term_start": "01/01/2026",
                    "term_end": "01/31/2026",
                    "list_unit_price_raw": "USD 0.00",
                    "discount_pct_raw": "0.00",
                    "net_unit_price_raw": "Included",
                    "net_total_raw": "USD 0.00",
                },
            ],
        }
    ]

    rows = _build_template_rows(files_payload=files_payload, euro_rate=1.17, margin_percent=20.0)

    assert len(rows) == 2
    assert rows[0]["Item"] == "SKU-A"
    assert rows[1]["Item"] == "SKU-B"
    assert rows[1]["Salesprice"] == 0.0
    assert rows[1]["Salesdiscount"] is None
    assert rows[0]["Date"] == "01/01/2026"
    assert rows[0]["Expires"] == "31/01/2026"
    assert rows[0]["ExpectedClose"] == "31/01/2026"
    assert rows[0]["ContractStart"] == "01/01/2026"
    assert rows[0]["ContractEnd"] == "31/01/2026"
    assert rows[0]["BusinessUnit"] == "Spain"
    assert rows[0]["Currency"] == "EUR"
    assert rows[0]["Location"] == "EXN Spain : ES Sales Stock"
    assert rows[0]["SalesCurrency"] == "EUR"
    assert rows[0]["Opportunity"] is None
    assert rows[0]["Quote ID (Line)"] == "Q-TEST"
