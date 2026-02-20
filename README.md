# Lossless PDF to Excel Extraction (Netskope Quotes)

This project extracts Netskope quote PDFs into:
- A multi-sheet Excel workbook for operations and review
- A JSON artifact for auditability and integrations

It is designed to avoid missing details by capturing:
- Document metadata and page stats
- Word-level and line-level text with coordinates
- Raw table cells (including continuation rows and empty cells)
- Parsed line items and business summary fields
- Hyperlinks and image metadata
- Validation/reconciliation results

## Install

```powershell
python -m pip install -r requirements.txt
```

## Frontend Integration

The backend now serves the frontend built from `pdf-to-excel-uploader`.

Build frontend assets:

```powershell
cd pdf-to-excel-uploader
npm ci
npm run build
cd ..
```

After build, the Flask app serves:
- `/` -> frontend app (`pdf-to-excel-uploader/dist/index.html`)
- `/api/*` -> backend API

Optional env var:
- `FRONTEND_DIST_DIR`: override frontend dist directory path

## Vercel Deployment

This project is ready for Vercel with `app.py` as the Python entrypoint.

- API endpoints:
  - `GET /health`
  - `POST /extract-template`
- API aliases (also supported): `GET /api/health`, `POST /api/extract-template`
- Serverless-safe behavior:
  - Uses temporary files only
  - No persistent local state required
  - OCR defaults to `off` (recommended on Vercel unless you provide native OCR binaries)

### Deploy

```powershell
vercel
```

Vercel config already runs frontend build during deploy:
- `cd pdf-to-excel-uploader && npm ci && npm run build`

### Optional Environment Variables

- `DEFAULT_TEMPLATE_PATH`: path to default template in deployment (fallback: `templates/Example with calculations.xlsx`)
- `CONFIG_PATH`: alternate config file path (default: `config.yaml`)
- `EXTRACTOR_MAX_CONTENT_MB`: upload limit in MB (default: `20`)
- `CORS_ALLOW_ORIGIN`: CORS origin (default: `*`)

### API Usage Example

`multipart/form-data` request with `pdf` and optionally `template`:

```bash
curl -X POST "<YOUR_URL>/extract-template" \
  -F "pdf=@Q-220053-20251224-0752  (1).pdf" \
  -F "template=@Example with calculations.xlsx" \
  -F "strict=true" \
  -F "template_only=true" \
  -F "ocr_mode=off" \
  -F "euro_rate=1.17" \
  -F "margin_percent=10"
```

If you want JSON response instead of file download, add:
- `return_json=true`
- `dedupe=true` to skip identical duplicate PDFs in the same request (default: `false`)

Send multiple PDFs in one request by repeating `pdf`:

```bash
curl -X POST "<YOUR_URL>/extract-template" \
  -F "pdf=@quote1.pdf" \
  -F "pdf=@quote2.pdf" \
  -F "template=@Example with calculations.xlsx" \
  -F "strict=true" \
  -F "template_only=true" \
  -F "ocr_mode=off" \
  -F "euro_rate=1.17" \
  -F "margin_percent=10"
```

By default, all uploaded files are processed, even if two files are identical.
To enable duplicate skipping, send:
- `dedupe=true`

File-download responses include summary headers:
- `X-Uploaded-Files`
- `X-Processed-Files`
- `X-Duplicates-Skipped`
- `X-Dedupe-Enabled`
- `X-Rows-Written`

## Windows OCR Prerequisites (Optional)

OCR is only needed for scanned/low-text pages.

1. Install Tesseract OCR and ensure `tesseract.exe` is on `PATH`.
2. Install Poppler and either:
- Add Poppler `bin` folder to `PATH`, or
- Pass `--poppler-path "C:\path\to\poppler\Library\bin"` at runtime.

## CLI

```powershell
python extract.py --input <file_or_folder> --output <output.xlsx> --json-output <output.json> --config config.yaml --ocr-mode auto --strict
```

### Main Options

- `--input`: one or more PDF files and/or directories
- `--output`: output Excel path (default: `output.xlsx`)
- `--json-output`: output JSON path (default: same as output with `.json`)
- `--config`: YAML config path
- `--ocr-mode auto|off|always`:
  - `auto` (default): OCR only low-text pages
  - `off`: no OCR
  - `always`: OCR every page
- `--strict / --no-strict`: strict validation enabled by default
- `--char-layer`: include optional char-level layer
- `--no-char-layer`: disable char-layer (default behavior)
- `--no-tables`: disable table extraction
- `--tesseract-cmd`: explicit path to `tesseract.exe`
- `--poppler-path`: explicit Poppler bin path
- `--template`: optional Excel template path (fills yellow columns)
- `--template-output`: output for filled template (default: `<output>.template_filled.xlsx`)
- `--template-sheet`: template sheet name (default: `QuoteExportResults`)
- `--template-header-row`: optional header row index (auto-detected if omitted)
- `--template-data-start-row`: optional first line row index (defaults to `header_row + 1`)
- `--template-only`: skip the audit workbook and generate only the filled template + JSON
- `--euro-rate`: required for template output, used in Salesdiscount formula
- `--margin-percent`: required for template output, used as margin % in Salesdiscount formula (`10` => multiplier `1.1`)

## Output Sheets

- `document_metadata`
- `pages`
- `text_lines`
- `text_words`
- `tables_raw`
- `line_items_parsed`
- `links`
- `images`
- `business_summary`
- `validation_report`
- `text_chars` (only when `--char-layer` is enabled)

## Config Schema (`config.yaml`)

Top-level keys:
- `field_patterns`
- `table_settings`
- `line_item_rules`
- `normalization`
- `validation`
- `ocr`

You can refine regexes and table rules for additional quote variants.

## Tests

Run tests:

```powershell
python -m pytest -q
```

Integration tests use:
- `tests/fixtures/Q-220053-20251224-0752  (1).pdf`

## Example

```powershell
python extract.py `
  --input "tests\\fixtures\\Q-220053-20251224-0752  (1).pdf" `
  --output "quote_output.xlsx" `
  --json-output "quote_output.json" `
  --config "config.yaml" `
  --template "tests\\fixtures\\Example with calculations.xlsx" `
  --template-output "quote_template_filled.xlsx" `
  --template-only `
  --euro-rate 1.17 `
  --margin-percent 10 `
  --ocr-mode auto `
  --strict
```

## Batch Example (Multiple PDFs)

```powershell
python extract.py `
  --input "C:\\quotes\\quote1.pdf" "C:\\quotes\\quote2.pdf" "C:\\quotes\\archive" `
  --config "config.yaml" `
  --template "C:\\templates\\Example with calculations.xlsx" `
  --template-output "quote_template_filled_batch.xlsx" `
  --template-only `
  --euro-rate 1.17 `
  --margin-percent 10 `
  --ocr-mode off `
  --strict
```
