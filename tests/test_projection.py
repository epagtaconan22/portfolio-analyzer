import pytest
from app.ui.projection import compute_prop_projection


def _kpi(month, actual=None, budget=None, year=2026):
    return {
        "year": year, "month": month,
        "actual_income": actual, "budget_income": budget,
        "actual_expenses": actual * 0.4 if actual else None,
        "budget_expenses": budget * 0.4 if budget else None,
        "actual_noi": actual * 0.6 if actual else None,
        "budget_noi": budget * 0.6 if budget else None,
    }


def test_returns_empty_when_no_kpis():
    label, data = compute_prop_projection([])
    assert label == ""
    assert data == {}


def test_uses_latest_year():
    kpis = [_kpi(1, 10000, 9500, year=2025), _kpi(1, 12000, 11000, year=2026)]
    label, _ = compute_prop_projection(kpis)
    assert label == "2026"


def test_projection_q1_plus_q2q4_budget():
    kpis = [
        _kpi(1, 10000, 9000), _kpi(2, 10000, 9000), _kpi(3, 10000, 9000),
        _kpi(4, None, 9000),  _kpi(5, None, 9000),  _kpi(6, None, 9000),
        _kpi(7, None, 9000),  _kpi(8, None, 9000),  _kpi(9, None, 9000),
        _kpi(10, None, 9000), _kpi(11, None, 9000), _kpi(12, None, 9000),
    ]
    label, proj = compute_prop_projection(kpis)
    ai = proj["actual_income"]
    # Q1 actual income = 30000; Q2-Q4 budget income = 9 × 9000 = 81000; total = 111000
    assert ai["proj_fy"] == pytest.approx(111000)
    assert ai["fy_budget"] == pytest.approx(108000)  # 12 × 9000


def test_fallback_when_no_q2q4_budget():
    # Only Q1 data present — falls back to Q1 budget × 3 for Q2-Q4
    kpis = [_kpi(1, 10000, 9000), _kpi(2, 10000, 9000), _kpi(3, 10000, 9000)]
    _, proj = compute_prop_projection(kpis)
    ai = proj["actual_income"]
    # Q1 actual = 30000; Q2-Q4 fallback = 9000 × 3 = 27000; proj_fy = 57000
    assert ai["proj_fy"] == pytest.approx(57000)
    # FY budget fallback = 9000 × 4 = 36000
    assert ai["fy_budget"] == pytest.approx(36000)
