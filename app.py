from __future__ import annotations

import mimetypes
import os
import hashlib
import tempfile
from pathlib import Path
from typing import Any

from flask import Flask, Response, jsonify, request
from werkzeug.utils import secure_filename

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
    try:
        parsed = float(value)
    except ValueError:
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
            "endpoints": ["/health", "/extract-template"],
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
            "endpoints": ["/health", "/extract-template"],
            "default_ocr_mode": "off",
        }
    )


@app.route("/health", methods=["GET"])
@app.route("/api/health", methods=["GET"])
def health() -> Response:
    return jsonify({"ok": True})


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

    config_path = _config_path()
    if not config_path.exists():
        return jsonify({"ok": False, "error": f"Config file not found: {config_path}"}), 500

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)
        seen_digests: set[str] = set()
        saved_pdf_count = 0
        for index, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
            pdf_name = secure_filename(uploaded_pdf.filename) or f"input_{index}.pdf"
            if not pdf_name.lower().endswith(".pdf"):
                pdf_name = f"{pdf_name}.pdf"
            pdf_path = temp_dir / f"{index:03d}_{pdf_name}"
            uploaded_pdf.save(pdf_path)
            digest = hashlib.sha256(pdf_path.read_bytes()).hexdigest()
            if digest in seen_digests:
                pdf_path.unlink(missing_ok=True)
                continue
            seen_digests.add(digest)
            saved_pdf_count += 1

        if saved_pdf_count == 0:
            return jsonify({"ok": False, "error": "All uploaded PDFs were duplicates."}), 400

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
        response.headers["X-Processed-Files"] = str(payload.get("file_count", 0))
        response.headers["X-Rows-Written"] = str(template_summary.get("rows_written", 0))
        return response


@app.route("/<path:route_path>", methods=["GET"])
def frontend_routes(route_path: str) -> Response:
    if route_path.startswith("api/"):
        return jsonify({"ok": False, "error": "Not Found"}), 404
    frontend = _frontend_response_for_path(route_path=route_path)
    if frontend is not None:
        return frontend
    return jsonify({"ok": False, "error": "Not Found"}), 404
