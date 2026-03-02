from __future__ import annotations

import mimetypes
import os
import hashlib
import re
import tempfile
from pathlib import Path
from typing import Any

from flask import Flask, Response, jsonify, request
from werkzeug.utils import secure_filename

from pdf_quote_extractor.normalize import parse_number_value
from pdf_quote_extractor.pipeline import run_pipeline


def _bool_from_str(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    return default


def _parse_required_float(value: str | None, field_name: str) -> tuple[float | None, str | None]:
    if value is None or not value.strip():
        return None, f"Missing required field: {field_name}"
    parsed = parse_number_value(value, allow_thousands=True)
    if parsed is None:
        return None, f"Invalid numeric value for {field_name}"
    return parsed, None


def _default_template_path() -> Path | None:
    configured = os.getenv("DEFAULT_TEMPLATE_PATH")
    if configured:
        candidate = Path(configured)
        if candidate.exists():
            return candidate
    bundled = Path("templates") / "Example with calculations.xlsx"
    if bundled.exists():
        return bundled
    return None


def _frontend_dist_dir() -> Path:
    configured = os.getenv("FRONTEND_DIST_DIR")
    if configured:
        return Path(configured)
    return Path("pdf-to-excel-uploader") / "dist"


def _frontend_index_path() -> Path:
    return _frontend_dist_dir() / "index.html"


def _frontend_response_for_path(route_path: str | None = None) -> Response | None:
    dist_dir = _frontend_dist_dir()
    if not dist_dir.exists():
        return None

    if route_path:
        candidate = (dist_dir / route_path).resolve()
        if candidate.exists() and candidate.is_file() and str(candidate).startswith(str(dist_dir.resolve())):
            mimetype = mimetypes.guess_type(candidate.name)[0] or "application/octet-stream"
            return Response(candidate.read_bytes(), mimetype=mimetype)

    index_path = _frontend_index_path()
    if not index_path.exists():
        return None
    return Response(index_path.read_bytes(), mimetype="text/html")


def _config_path() -> Path:
    configured = os.getenv("CONFIG_PATH")
    if configured:
        return Path(configured)
    return Path("config.yaml")


def _normalize_vendor_key(raw: str | None) -> str:
    if raw is None:
        return ""
    normalized = re.sub(r"[^a-z0-9]+", "_", raw.strip().lower()).strip("_")
    return normalized


def _default_vendor() -> str:
    configured = _normalize_vendor_key(os.getenv("DEFAULT_VENDOR"))
    return configured or "netskope"


def _vendor_config_dir() -> Path:
    configured = os.getenv("VENDOR_CONFIG_DIR")
    if configured:
        return Path(configured)
    return Path("config") / "vendors"


def _resolve_config_path(vendor_raw: str | None) -> tuple[str, Path]:
    default_vendor = _default_vendor()
    requested_vendor = _normalize_vendor_key(vendor_raw) or default_vendor

    env_override = os.getenv(f"CONFIG_PATH_{requested_vendor.upper()}")
    if env_override:
        return requested_vendor, Path(env_override)

    vendor_dir = _vendor_config_dir()
    for extension in ("yaml", "yml"):
        candidate = vendor_dir / f"{requested_vendor}.{extension}"
        if candidate.exists():
            return requested_vendor, candidate

    if requested_vendor == default_vendor:
        return requested_vendor, _config_path()

    raise ValueError(
        f"Unsupported vendor: {requested_vendor}. Add a config at "
        f"{vendor_dir / f'{requested_vendor}.yaml'} or set CONFIG_PATH_{requested_vendor.upper()}."
    )


def _discover_vendors() -> list[str]:
    default_vendor = _default_vendor()
    vendors: list[str] = [default_vendor]
    seen = {default_vendor}

    configured = os.getenv("AVAILABLE_VENDORS", "")
    for entry in configured.split(","):
        vendor = _normalize_vendor_key(entry)
        if vendor and vendor not in seen:
            seen.add(vendor)
            vendors.append(vendor)

    for env_key in os.environ:
        if not env_key.startswith("CONFIG_PATH_"):
            continue
        vendor = _normalize_vendor_key(env_key.replace("CONFIG_PATH_", "", 1))
        if vendor and vendor not in seen:
            seen.add(vendor)
            vendors.append(vendor)

    vendor_dir = _vendor_config_dir()
    if vendor_dir.exists():
        for candidate in sorted(vendor_dir.glob("*.y*ml")):
            vendor = _normalize_vendor_key(candidate.stem)
            if vendor and vendor not in seen:
                seen.add(vendor)
                vendors.append(vendor)
    return vendors


def _vendor_label(vendor: str) -> str:
    words = [word for word in vendor.replace("_", " ").split() if word]
    if not words:
        return "Vendor"
    return " ".join(word.capitalize() for word in words)


def _cors_origin() -> str:
    return os.getenv("CORS_ALLOW_ORIGIN", "*")


app = Flask(__name__)
max_content_mb = int(os.getenv("EXTRACTOR_MAX_CONTENT_MB", "20"))
app.config["MAX_CONTENT_LENGTH"] = max_content_mb * 1024 * 1024


@app.after_request
def add_cors_headers(resp: Response) -> Response:
    resp.headers["Access-Control-Allow-Origin"] = _cors_origin()
    resp.headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type,Authorization"
    return resp


@app.route("/", methods=["GET"])
def root() -> Response:
    frontend = _frontend_response_for_path()
    if frontend is not None:
        return frontend
    return jsonify(
        {
            "ok": True,
            "service": "pdf-template-extractor",
            "version": "1.0",
            "endpoints": ["/health", "/vendors", "/extract-template"],
            "default_ocr_mode": "off",
        }
    )


@app.route("/api", methods=["GET"])
def api_root() -> Response:
    return jsonify(
        {
            "ok": True,
            "service": "pdf-template-extractor",
            "version": "1.0",
            "endpoints": ["/health", "/vendors", "/extract-template"],
            "default_ocr_mode": "off",
        }
    )


@app.route("/health", methods=["GET"])
@app.route("/api/health", methods=["GET"])
def health() -> Response:
    return jsonify({"ok": True})


@app.route("/vendors", methods=["GET"])
@app.route("/api/vendors", methods=["GET"])
def vendors() -> Response:
    available = _discover_vendors()
    default_vendor = _default_vendor()
    return jsonify(
        {
            "ok": True,
            "default_vendor": default_vendor,
            "vendors": [{"id": vendor, "label": _vendor_label(vendor)} for vendor in available],
        }
    )


@app.route("/extract-template", methods=["POST", "OPTIONS"])
@app.route("/api/extract-template", methods=["POST", "OPTIONS"])
def extract_template() -> Response:
    if request.method == "OPTIONS":
        return Response(status=204)

    uploaded_pdfs = [
        file for file in (request.files.getlist("pdf") + request.files.getlist("pdfs")) if file and file.filename
    ]
    if not uploaded_pdfs:
        return jsonify({"ok": False, "error": "Missing required file field: pdf (or pdfs)"}), 400

    uploaded_template = request.files.get("template")
    default_template = _default_template_path()
    if uploaded_template is None and default_template is None:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": (
                        "No template uploaded and no default template found. "
                        "Provide form-data file field 'template' or set DEFAULT_TEMPLATE_PATH."
                    ),
                }
            ),
            400,
        )

    ocr_mode = request.form.get("ocr_mode", "off")
    if ocr_mode not in {"off", "auto", "always"}:
        return jsonify({"ok": False, "error": "ocr_mode must be one of: off, auto, always"}), 400

    strict = _bool_from_str(request.form.get("strict"), default=True)
    template_only = _bool_from_str(request.form.get("template_only"), default=True)
    return_json = _bool_from_str(request.form.get("return_json"), default=False)
    dedupe_uploads = _bool_from_str(request.form.get("dedupe"), default=False)
    euro_rate, euro_rate_error = _parse_required_float(request.form.get("euro_rate"), "euro_rate")
    if euro_rate_error:
        return jsonify({"ok": False, "error": euro_rate_error}), 400
    if euro_rate is not None and euro_rate <= 0:
        return jsonify({"ok": False, "error": "euro_rate must be greater than 0"}), 400

    margin_percent, margin_error = _parse_required_float(
        request.form.get("margin_percent"),
        "margin_percent",
    )
    if margin_error:
        return jsonify({"ok": False, "error": margin_error}), 400

    vendor_raw = request.form.get("vendor")
    try:
        selected_vendor, config_path = _resolve_config_path(vendor_raw)
    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    if not config_path.exists():
        return jsonify({"ok": False, "error": f"Config file not found: {config_path}"}), 500

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)
        seen_digests: set[str] = set()
        saved_pdf_count = 0
        duplicates_skipped = 0
        for index, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
            pdf_name = secure_filename(uploaded_pdf.filename) or f"input_{index}.pdf"
            if not pdf_name.lower().endswith(".pdf"):
                pdf_name = f"{pdf_name}.pdf"
            pdf_path = temp_dir / f"{index:03d}_{pdf_name}"
            uploaded_pdf.save(pdf_path)
            if dedupe_uploads:
                digest = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
                if digest in seen_digests:
                    duplicates_skipped += 1
                    pdf_path.unlink(missing_ok=True)
                    continue
                seen_digests.add(digest)
            saved_pdf_count += 1

        if saved_pdf_count == 0:
            if dedupe_uploads and duplicates_skipped > 0:
                return (
                    jsonify(
                        {
                            "ok": False,
                            "error": (
                                "All uploaded PDFs were duplicates with dedupe enabled. "
                                "Set dedupe=false to process all uploads."
                            ),
                        }
                    ),
                    400,
                )
            return jsonify({"ok": False, "error": "No valid PDFs were processed."}), 400

        if uploaded_template is not None and uploaded_template.filename:
            template_name = secure_filename(uploaded_template.filename) or "template.xlsx"
            if not template_name.lower().endswith(".xlsx"):
                template_name = f"{template_name}.xlsx"
            template_path = temp_dir / template_name
            uploaded_template.save(template_path)
        else:
            template_path = default_template
            assert template_path is not None

        audit_output_path = temp_dir / "audit_output.xlsx"
        json_output_path = temp_dir / "run_output.json"
        template_output_path = temp_dir / "filled_template.xlsx"

        try:
            exit_code, payload = run_pipeline(
                input_path=temp_dir,
                output_path=audit_output_path,
                json_output_path=json_output_path,
                config_path=config_path,
                ocr_mode=ocr_mode,
                strict=strict,
                include_char_layer=False,
                include_tables=True,
                tesseract_cmd=None,
                poppler_path=None,
                template_path=template_path,
                template_output_path=template_output_path,
                euro_rate=euro_rate,
                margin_percent=margin_percent,
                write_audit_workbook=not template_only,
            )
        except Exception as exc:
            return jsonify({"ok": False, "error": f"Extraction failed: {exc}"}), 500

        payload["upload_summary"] = {
            "uploaded_files": len(uploaded_pdfs),
            "processed_files": saved_pdf_count,
            "duplicates_skipped": duplicates_skipped,
            "dedupe_enabled": dedupe_uploads,
        }
        payload["vendor"] = selected_vendor

        if strict and exit_code != 0:
            return jsonify({"ok": False, "error": "Strict validation failed.", "payload": payload}), 422

        if return_json:
            return jsonify({"ok": True, "payload": payload})

        file_bytes = template_output_path.read_bytes()
        response = Response(
            file_bytes,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers["Content-Disposition"] = 'attachment; filename="filled_template.xlsx"'
        template_summary: dict[str, Any] = payload.get("template_output", {}) if isinstance(payload, dict) else {}
        upload_summary: dict[str, Any] = payload.get("upload_summary", {}) if isinstance(payload, dict) else {}
        response.headers["X-Processed-Files"] = str(payload.get("file_count", 0))
        response.headers["X-Uploaded-Files"] = str(upload_summary.get("uploaded_files", 0))
        response.headers["X-Duplicates-Skipped"] = str(upload_summary.get("duplicates_skipped", 0))
        response.headers["X-Dedupe-Enabled"] = str(upload_summary.get("dedupe_enabled", False)).lower()
        response.headers["X-Rows-Written"] = str(template_summary.get("rows_written", 0))
        response.headers["X-Vendor"] = selected_vendor
        return response


@app.route("/<path:route_path>", methods=["GET"])
def frontend_routes(route_path: str) -> Response:
    if route_path.startswith("api/"):
        return jsonify({"ok": False, "error": "Not Found"}), 404
    frontend = _frontend_response_for_path(route_path=route_path)
    if frontend is not None:
        return frontend
    return jsonify({"ok": False, "error": "Not Found"}), 404
