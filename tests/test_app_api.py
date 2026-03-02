import io
from pathlib import Path

from app import app
from werkzeug.datastructures import MultiDict


def test_health_endpoint():
    client = app.test_client()
    response = client.get("/health")
    assert response.status_code == 200
    assert response.get_json() == {"ok": True}


def test_vendors_endpoint():
    client = app.test_client()
    response = client.get("/api/vendors")
    assert response.status_code == 200
    body = response.get_json()
    assert body["ok"] is True
    assert body["default_vendor"] == "netskope"
    assert any(vendor["id"] == "netskope" for vendor in body["vendors"])


def test_extract_template_return_json():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = {
        "pdf": (io.BytesIO(pdf_path.read_bytes()), "quote.pdf"),
        "template": (io.BytesIO(template_path.read_bytes()), "template.xlsx"),
        "strict": "true",
        "template_only": "true",
        "ocr_mode": "off",
        "euro_rate": "1.17",
        "margin_percent": "10",
        "return_json": "true",
    }
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["ok"] is True
    assert body["payload"]["file_count"] == 1
    assert body["payload"]["template_output"]["rows_written"] == 4


def test_extract_template_return_file():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = {
        "pdf": (io.BytesIO(pdf_path.read_bytes()), "quote.pdf"),
        "template": (io.BytesIO(template_path.read_bytes()), "template.xlsx"),
        "strict": "true",
        "template_only": "true",
        "ocr_mode": "off",
        "euro_rate": "1.17",
        "margin_percent": "10",
    }
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    assert (
        response.headers["Content-Type"]
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert response.headers["X-Rows-Written"] == "4"
    assert len(response.data) > 1000


def test_extract_template_with_multiple_identical_pdfs_keeps_all_by_default():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = MultiDict(
        [
            ("pdf", (io.BytesIO(pdf_path.read_bytes()), "quote_1.pdf")),
            ("pdf", (io.BytesIO(pdf_path.read_bytes()), "quote_2.pdf")),
            ("template", (io.BytesIO(template_path.read_bytes()), "template.xlsx")),
            ("strict", "true"),
            ("template_only", "true"),
            ("ocr_mode", "off"),
            ("euro_rate", "1.17"),
            ("margin_percent", "10"),
            ("return_json", "true"),
        ]
    )
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["ok"] is True
    assert body["payload"]["file_count"] == 2
    assert body["payload"]["template_output"]["rows_written"] == 8
    assert body["payload"]["upload_summary"]["duplicates_skipped"] == 0


def test_extract_template_with_multiple_pdfs_deduplicates_when_enabled():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = MultiDict(
        [
            ("pdf", (io.BytesIO(pdf_path.read_bytes()), "quote_1.pdf")),
            ("pdf", (io.BytesIO(pdf_path.read_bytes()), "quote_2.pdf")),
            ("template", (io.BytesIO(template_path.read_bytes()), "template.xlsx")),
            ("strict", "true"),
            ("template_only", "true"),
            ("ocr_mode", "off"),
            ("euro_rate", "1.17"),
            ("margin_percent", "10"),
            ("return_json", "true"),
            ("dedupe", "true"),
        ]
    )
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["ok"] is True
    assert body["payload"]["file_count"] == 1
    assert body["payload"]["template_output"]["rows_written"] == 4
    assert body["payload"]["upload_summary"]["duplicates_skipped"] == 1


def test_extract_template_missing_euro_rate_returns_400():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = {
        "pdf": (io.BytesIO(pdf_path.read_bytes()), "quote.pdf"),
        "template": (io.BytesIO(template_path.read_bytes()), "template.xlsx"),
        "strict": "true",
        "template_only": "true",
        "ocr_mode": "off",
        "margin_percent": "10",
    }
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 400
    body = response.get_json()
    assert body["ok"] is False
    assert "euro_rate" in body["error"]


def test_extract_template_accepts_comma_decimal_and_vendor():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = {
        "pdf": (io.BytesIO(pdf_path.read_bytes()), "quote.pdf"),
        "template": (io.BytesIO(template_path.read_bytes()), "template.xlsx"),
        "strict": "true",
        "template_only": "true",
        "ocr_mode": "off",
        "euro_rate": "1,17",
        "margin_percent": "10,5",
        "vendor": "netskope",
        "return_json": "true",
    }
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["ok"] is True
    assert body["payload"]["vendor"] == "netskope"
    assert body["payload"]["template_output"]["rows_written"] == 4


def test_extract_template_rejects_unknown_vendor():
    client = app.test_client()
    project_root = Path(__file__).resolve().parents[1]
    pdf_path = project_root / "tests" / "fixtures" / "Q-220053-20251224-0752  (1).pdf"
    template_path = project_root / "tests" / "fixtures" / "Example with calculations.xlsx"

    payload = {
        "pdf": (io.BytesIO(pdf_path.read_bytes()), "quote.pdf"),
        "template": (io.BytesIO(template_path.read_bytes()), "template.xlsx"),
        "strict": "true",
        "template_only": "true",
        "ocr_mode": "off",
        "euro_rate": "1.17",
        "margin_percent": "10",
        "vendor": "unknown-vendor",
    }
    response = client.post("/extract-template", data=payload, content_type="multipart/form-data")
    assert response.status_code == 400
    body = response.get_json()
    assert body["ok"] is False
    assert "Unsupported vendor" in body["error"]


def test_root_serves_frontend_when_dist_exists(tmp_path, monkeypatch):
    dist_dir = tmp_path / "dist"
    dist_dir.mkdir(parents=True, exist_ok=True)
    (dist_dir / "index.html").write_text("<html><body>frontend-ok</body></html>", encoding="utf-8")
    monkeypatch.setenv("FRONTEND_DIST_DIR", str(dist_dir))

    client = app.test_client()
    response = client.get("/")
    assert response.status_code == 200
    assert "text/html" in response.headers["Content-Type"]
    assert b"frontend-ok" in response.data


def test_spa_fallback_route_serves_index(tmp_path, monkeypatch):
    dist_dir = tmp_path / "dist"
    dist_dir.mkdir(parents=True, exist_ok=True)
    (dist_dir / "index.html").write_text("<html><body>spa-fallback</body></html>", encoding="utf-8")
    monkeypatch.setenv("FRONTEND_DIST_DIR", str(dist_dir))

    client = app.test_client()
    response = client.get("/dashboard")
    assert response.status_code == 200
    assert "text/html" in response.headers["Content-Type"]
    assert b"spa-fallback" in response.data
