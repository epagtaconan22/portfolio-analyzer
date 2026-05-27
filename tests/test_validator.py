import os
import openpyxl
import pytest
from app.exporter.validator import validate_workbook, validate_both_workbooks
from app.models import QualityCheck


def _make_workbook(tabs: list[str], tmp_path) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for tab in tabs:
        wb.create_sheet(tab)
    path = str(tmp_path / "test.xlsx")
    wb.save(path)
    return path


def test_valid_zip(tmp_path):
    path = _make_workbook(["Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs"], tmp_path)
    checks = validate_workbook(path, {"Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs"})
    zip_check = next(c for c in checks if "Valid ZIP" in c.check_name)
    assert zip_check.passed


def test_missing_tabs_flagged(tmp_path):
    path = _make_workbook(["Dashboard"], tmp_path)
    checks = validate_workbook(path, {"Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs"})
    tab_check = next(c for c in checks if "tabs" in c.check_name.lower())
    assert not tab_check.passed
    assert "Property Quarterly KPIs" in tab_check.detail


def test_all_tabs_present(tmp_path):
    path = _make_workbook(["Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs"], tmp_path)
    checks = validate_workbook(path, {"Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs"})
    tab_check = next(c for c in checks if "tabs" in c.check_name.lower())
    assert tab_check.passed


def test_no_formula_errors(tmp_path):
    path = _make_workbook(["Sheet1"], tmp_path)
    checks = validate_workbook(path, {"Sheet1"})
    err_check = next(c for c in checks if "formula" in c.check_name.lower())
    assert err_check.passed
    assert err_check.detail == "Clean"


def test_backup_workbook_smoke(tmp_path):
    """Smoke test: build_backup_workbook produces a valid file that passes validation."""
    from app.models import (
        MappedRow, PropertyPeriodKPIs, SourceIndexEntry, MappingEntry, QualityCheck
    )
    from app.exporter.backup_workbook import build_backup_workbook

    row = MappedRow(
        property_name="Test Prop", pm_name="PM Co",
        source_workbook="wb.xlsx", source_sheet="S1", source_type="Actual",
        source_row=1, account_code="", account_name="Gross Potential Rent",
        year=2024, month=1, amount=10000, original_amount=10000, notes="",
        account_category="Rental Income", kpi_mapping="GPR / Rental Income",
        include_in_noi=True, include_in_eco_occ=True, treatment="Income",
    )
    k = PropertyPeriodKPIs("Test Prop", "PM Co", 2024, 1, "Jan")
    k.actual_income = 10000; k.actual_noi = 6000

    src = SourceIndexEntry(
        source_workbook="wb.xlsx", source_sheet="S1",
        property_name="Test Prop", pm_name="PM Co",
        year=2024, source_type="Actual", processed=True, rows_extracted=1,
    )
    mapping = MappingEntry(
        account_code="", account_name="Gross Potential Rent",
        assigned_category="Rental Income", kpi_mapping="GPR / Rental Income",
        treatment="Income", include_in_noi=True, include_in_eco_occ=True,
    )
    qc = QualityCheck("Test check", True, "OK")

    path = str(tmp_path / "backup.xlsx")
    build_backup_workbook([row], [k], [src], [mapping], [qc], path)

    from app.exporter.validator import validate_workbook, _BACKUP_TABS
    checks = validate_workbook(path, _BACKUP_TABS)
    zip_check = next(c for c in checks if "Valid ZIP" in c.check_name)
    tab_check = next(c for c in checks if "tabs" in c.check_name.lower())
    assert zip_check.passed
    assert tab_check.passed
