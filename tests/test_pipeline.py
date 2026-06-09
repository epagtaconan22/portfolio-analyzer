"""Integration tests for the analysis pipeline — replaces test_routes.py."""
import io
import os
import pytest
import openpyxl

from app.ui.pipeline import run_analysis_pipeline


def _make_fin_workbook_bytes(sheet_title="Actual - Test Property"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
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


@pytest.fixture(autouse=True)
def chdir_tmp(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("runs", exist_ok=True)


def test_pipeline_returns_run_id():
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[],
        ar_files=[],
        settings={"portfolio_name": "Test Portfolio", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": ["PM One"],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    assert run_id
    assert os.path.isdir(os.path.join("runs", run_id))


def test_pipeline_writes_metadata():
    import json
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "My Portfolio", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    meta_path = os.path.join("runs", run_id, "metadata.json")
    assert os.path.isfile(meta_path)
    with open(meta_path) as f:
        meta = json.load(f)
    assert meta["portfolio_name"] == "My Portfolio"


def test_pipeline_writes_main_workbook():
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "WB Test", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    run_dir = os.path.join("runs", run_id)
    xlsx_files = [f for f in os.listdir(run_dir) if f.endswith(".xlsx")]
    assert len(xlsx_files) == 2


def test_pipeline_no_files_raises():
    with pytest.raises(ValueError, match="No valid"):
        run_analysis_pipeline(
            fin_files=[], occ_files=[], ar_files=[],
            settings={"portfolio_name": "Empty", "eco_occ_target": 0.95,
                      "use_budget_eco_occ": False, "pm_names": [],
                      "excluded_properties": set(), "carveout_properties": set(),
                      "stabilized_properties": set(), "period_filter": "Full Year",
                      "selected_months": [], "custom_mapping": None},
        )


def test_pipeline_carveout_property_flagged():
    from app.storage.runs import load_run
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "CO Test", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(),
                  "carveout_properties": {"test property"},
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    data = load_run(run_id)
    carveout_kpis = [k for k in data["kpis"] if k.get("is_carveout")]
    assert len(carveout_kpis) > 0
