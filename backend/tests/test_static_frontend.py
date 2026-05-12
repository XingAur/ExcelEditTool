from pathlib import Path

from fastapi.testclient import TestClient

from backend.app.main import create_app


def test_create_app_serves_frontend_index(tmp_path: Path) -> None:
    (tmp_path / "index.html").write_text(
        "<!doctype html><title>ExcelEditTool</title>",
        encoding="utf-8",
    )

    client = TestClient(create_app(static_dir=tmp_path))
    response = client.get("/")

    assert response.status_code == 200
    assert "ExcelEditTool" in response.text
