import pytest
from app.models import OccupancyRow, PropertyPeriodKPIs
from app.calculator.physical_occ import enrich_physical_occ

def _occ(prop, month, occupied, total, year=2024):
    return OccupancyRow(property_name=prop, year=year, month=month,
                        occupied_units=occupied, total_units=total)

def _kpi(prop, month, year=2024, actual_income=10000, actual_expenses=4000, actual_noi=6000):
    k = PropertyPeriodKPIs(property_name=prop, pm_name="PM", year=year, month=month, period="Jan")
    k.actual_income = actual_income
    k.actual_expenses = actual_expenses
    k.actual_noi = actual_noi
    return k

def test_physical_occ_pct():
    occ = [_occ("Prop A", 1, 90, 100)]
    kpis = [_kpi("Prop A", 1)]
    result = enrich_physical_occ(occ, kpis)
    k = result[0]
    assert k.total_units == 100
    assert k.occupied_units == 90
    assert k.physical_occ_pct == pytest.approx(0.90)

def test_leakage_gap():
    occ = [_occ("Prop A", 1, 95, 100)]
    kpis = [_kpi("Prop A", 1)]
    kpis[0].eco_occ_pct = 0.88
    result = enrich_physical_occ(occ, kpis)
    # Leakage = physical - eco = 0.95 - 0.88 = 0.07
    assert result[0].leakage_gap == pytest.approx(0.07)

def test_per_unit_calculations():
    occ = [_occ("Prop A", 1, 95, 100)]
    kpis = [_kpi("Prop A", 1, actual_income=10000, actual_expenses=4000, actual_noi=6000)]
    result = enrich_physical_occ(occ, kpis)
    k = result[0]
    assert k.income_per_unit == pytest.approx(100.0)   # 10000 / 100
    assert k.expense_per_unit == pytest.approx(40.0)   # 4000 / 100
    assert k.noi_per_unit == pytest.approx(60.0)       # 6000 / 100

def test_per_unit_none_when_no_occupancy_report():
    kpis = [_kpi("Prop A", 1)]
    result = enrich_physical_occ([], kpis)
    assert result[0].income_per_unit is None
    assert result[0].total_units is None

def test_yoy_variance():
    occ_2023 = [_occ("Prop A", 1, 88, 100, year=2023)]
    occ_2024 = [_occ("Prop A", 1, 95, 100, year=2024)]
    kpis_2023 = [_kpi("Prop A", 1, year=2023)]
    kpis_2023[0].eco_occ_pct = 0.85
    kpis_2024 = [_kpi("Prop A", 1, year=2024)]
    kpis_2024[0].eco_occ_pct = 0.90
    all_kpis = enrich_physical_occ(occ_2023 + occ_2024, kpis_2023 + kpis_2024)
    k2024 = next(k for k in all_kpis if k.year == 2024)
    assert k2024.yoy_physical_occ_variance == pytest.approx(0.07)  # 0.95 - 0.88
    assert k2024.yoy_eco_occ_variance == pytest.approx(0.05)        # 0.90 - 0.85
