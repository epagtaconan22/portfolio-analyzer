"""Tests for the Yardi AR Aging file parser."""

import pytest
import openpyxl

from app.models import ARAgingRow
from app.parser.ar_aging import parse_ar_aging_reports


def _make_ar_wb(tmp_path, filename, rows):
    """Helper: create a synthetic Yardi AR Aging workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report1"
    ws.append(["Affordable Aging Detail"])
    ws.append(["Property: Affirmed Property List (affirmed)"])
    period = f"Post To(MM/YY): {filename.split('_')[-2].zfill(2)}/{filename.split('_')[-1].split('.')[0]}"
    ws.append([period])
    ws.append(["Property Name", "Charge Amount", "Current Owed",
               "0-30 Owed", "31-60 Owed", "61-90 Owed",
               "Over 90 Owed", "Pre-payments", "Suspense"])
    ws.append([""] * 9)
    for r in rows:
        ws.append(r)
    ws.append([None] * 9)
    ws.append(["Grand Total"] + [0] * 8)
    path = tmp_path / filename
    wb.save(str(path))
    return str(path)


@pytest.fixture
def tenant_rent_wb(tmp_path):
    return _make_ar_wb(tmp_path, "Solari_AR Aging_Tenant Rent_03_2024.xlsx", [
        ["Alora Family (alora)",  10000, 8000, 500, 300, 200, 100, -50,   0],
        ["Beechwood (beech)",     20000, 15000, 1000, 800, 600, 400, -100, 0],
    ])


@pytest.fixture
def subsidy_wb(tmp_path):
    return _make_ar_wb(tmp_path, "Solari_AR Aging_Subsidy_03_2024.xlsx", [
        ["Alora Family (alora)", 5000, 4000, 200, 100, 50, 25, 0, 0],
    ])


@pytest.fixture
def tenant_receivable_wb(tmp_path):
    return _make_ar_wb(tmp_path, "ConAm_AR Aging_Tenant Receivable_06_2024.xlsx", [
        ["Alora Family (alora)", 10000, 8000, 500, 300, 200, 100, -50, 0],
    ])


@pytest.fixture
def subsidy_receivable_wb(tmp_path):
    return _make_ar_wb(tmp_path, "ConAm_AR Aging_Subsidy Receivable_06_2024.xlsx", [
        ["Alora Family (alora)", 5000, 4000, 200, 100, 50, 25, 0, 0],
    ])


def test_returns_ar_aging_row_instances(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert len(rows) == 2
    assert all(isinstance(r, ARAgingRow) for r in rows)


def test_receivable_type_tenant_rent(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.receivable_type == "Tenant Rent" for r in rows)


def test_receivable_type_subsidy(subsidy_wb):
    rows = parse_ar_aging_reports([subsidy_wb])
    assert all(r.receivable_type == "Subsidy" for r in rows)


def test_type_normalization_tenant_receivable(tenant_receivable_wb):
    rows = parse_ar_aging_reports([tenant_receivable_wb])
    assert all(r.receivable_type == "Tenant Rent" for r in rows)


def test_type_normalization_subsidy_receivable(subsidy_receivable_wb):
    rows = parse_ar_aging_reports([subsidy_receivable_wb])
    assert all(r.receivable_type == "Subsidy" for r in rows)


def test_pm_name_from_filename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.pm_name == "Solari" for r in rows)


def test_year_month_from_filename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.year == 2024 and r.month == 3 for r in rows)


def test_property_code_suffix_stripped(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    prop_names = {r.property_name for r in rows}
    assert "Alora Family" in prop_names
    assert "Beechwood" in prop_names
    assert not any("(" in name for name in prop_names)


def test_numeric_fields(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    assert alora.charge_amount  == pytest.approx(10000)
    assert alora.current_owed   == pytest.approx(8000)
    assert alora.owed_0_30      == pytest.approx(500)
    assert alora.owed_31_60     == pytest.approx(300)
    assert alora.owed_61_90     == pytest.approx(200)
    assert alora.owed_over_90   == pytest.approx(100)
    assert alora.prepayments    == pytest.approx(-50)


def test_grand_total_row_excluded(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert not any("Grand Total" in r.property_name for r in rows)


def test_blank_row_stops_iteration(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.property_name for r in rows)


def test_total_overdue_computed(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    # total_over_60 = owed_61_90 + owed_over_90 = 200 + 100 = 300 (excludes 31-60 bucket)
    assert alora.total_overdue == pytest.approx(300)


def test_pct_overdue_computed(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    # pct_over_60 = 300 / 10000 = 0.03 (only 61+ days past due)
    assert alora.pct_overdue == pytest.approx(0.03)


def test_none_numeric_treated_as_zero(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report1"
    ws.append(["Affordable Aging Detail"])
    ws.append(["Property: Affirmed Property List (affirmed)"])
    ws.append(["Post To(MM/YY): 03/2024"])
    ws.append(["Property Name", "Charge Amount", "Current Owed",
               "0-30 Owed", "31-60 Owed", "61-90 Owed",
               "Over 90 Owed", "Pre-payments", "Suspense"])
    ws.append([""] * 9)
    ws.append(["Test Prop (test)", 5000, 5000, None, None, None, None, None, 0])
    ws.append([None] * 9)
    ws.append(["Grand Total"] + [0] * 8)
    path = tmp_path / "PM_AR Aging_Subsidy_03_2024.xlsx"
    wb.save(str(path))
    rows = parse_ar_aging_reports([str(path)])
    assert len(rows) == 1
    assert rows[0].owed_31_60   == 0.0
    assert rows[0].prepayments  == 0.0
    assert rows[0].pct_overdue  == 0.0


def test_source_file_is_basename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.source_file == "Solari_AR Aging_Tenant Rent_03_2024.xlsx" for r in rows)


def test_multiple_files_combined(tenant_rent_wb, subsidy_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb, subsidy_wb])
    types = {r.receivable_type for r in rows}
    assert types == {"Tenant Rent", "Subsidy"}
