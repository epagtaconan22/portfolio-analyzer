import io
import os
import openpyxl
import pytest
from app import create_app


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("runs", exist_ok=True)
    app = create_app()
    app.config["TESTING"] = True
    return app.test_client()


def _make_workbook_bytes():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Actual - Test Property"
    ws.append(["Account", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    ws.append(["Gross Potential Rent",
               5000, 5000, 5000, 5000, 5000, 5000,
               5000, 5000, 5000, 5000, 5000, 5000])
    ws.append(["Vacancy Loss",
               -250, -250, -250, -250, -250, -250,
               -250, -250, -250, -250, -250, -250])
    ws.append(["Management Fee",
               500, 500, 500, 500, 500, 500,
               500, 500, 500, 500, 500, 500])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def test_index_loads(client):
    resp = client.get("/")
    assert resp.status_code == 200
    assert b"Portfolio" in resp.data


def test_history_empty(client):
    resp = client.get("/history")
    assert resp.status_code == 200
    assert b"No analyses" in resp.data


def test_run_analysis_and_redirect(client):
    wb_bytes = _make_workbook_bytes()
    data = {
        "portfolio_name": "Test Portfolio",
        "eco_occ_target": "95",
        "pm_names": "PM One",
        "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
    }
    resp = client.post("/", data=data, content_type="multipart/form-data",
                       follow_redirects=False)
    assert resp.status_code == 302
    assert "/results/" in resp.headers["Location"]


def test_run_then_history_shows_portfolio(client):
    wb_bytes = _make_workbook_bytes()
    data = {
        "portfolio_name": "Test Portfolio",
        "eco_occ_target": "95",
        "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
    }
    client.post("/", data=data, content_type="multipart/form-data",
                follow_redirects=True)
    resp = client.get("/history")
    assert resp.status_code == 200
    assert b"Test Portfolio" in resp.data


def test_dashboard_loads_after_run(client):
    wb_bytes = _make_workbook_bytes()
    data = {
        "portfolio_name": "Test Portfolio",
        "eco_occ_target": "95",
        "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
    }
    resp = client.post("/", data=data, content_type="multipart/form-data",
                       follow_redirects=True)
    assert resp.status_code == 200
    assert b"Test Portfolio" in resp.data


def test_no_files_uploaded_shows_flash(client):
    data = {"portfolio_name": "Empty"}
    resp = client.post("/", data=data, content_type="multipart/form-data",
                       follow_redirects=True)
    assert resp.status_code == 200
    assert b"upload" in resp.data.lower()


def test_download_endpoint(client):
    wb_bytes = _make_workbook_bytes()
    data = {
        "portfolio_name": "Test Portfolio",
        "eco_occ_target": "95",
        "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
    }
    resp = client.post("/", data=data, content_type="multipart/form-data",
                       follow_redirects=False)
    location = resp.headers["Location"]
    run_id = location.rstrip("/").split("/")[-1]

    resp2 = client.get(f"/results/{run_id}/download")
    assert resp2.status_code == 200
    assert resp2.content_type == "application/zip"


def test_invalid_run_id_404(client):
    resp = client.get("/results/nonexistent_run_id_xyz")
    assert resp.status_code == 404
