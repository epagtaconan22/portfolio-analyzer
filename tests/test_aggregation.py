import pytest
from app.ui.aggregation import agg_kpis, month_to_quarter, quarter_label, quarter_months


def _kpi(income, expenses, gpr=None, vacancy=0, year=2025, month=1):
    noi = income - expenses if income is not None and expenses is not None else None
    net_coll = (gpr - vacancy) if gpr is not None else None
    return {
        "actual_income": income,
        "actual_expenses": expenses,
        "actual_noi": noi,
        "budget_income": income * 0.95 if income else None,
        "budget_expenses": expenses * 1.05 if expenses else None,
        "budget_noi": None,
        "gpr": gpr,
        "vacancy": vacancy,
        "concessions": 0,
        "bad_debt": 0,
        "net_collectible": net_coll,
        "eco_occ_pct": (net_coll / gpr) if (net_coll and gpr) else None,
        "budget_eco_occ_pct": None,
        "physical_occ_pct": None,
        "occupied_units": None,
        "total_units": 100,
        "income_per_unit": None,
        "expense_per_unit": None,
        "noi_per_unit": None,
        "year": year,
        "month": month,
    }


def test_agg_kpis_sums_income():
    kpis = [_kpi(10000, 4000), _kpi(8000, 3000)]
    result = agg_kpis(kpis)
    assert result["actual_income"] == pytest.approx(18000)


def test_agg_kpis_computes_noi():
    kpis = [_kpi(10000, 4000)]
    result = agg_kpis(kpis)
    assert result["actual_noi"] == pytest.approx(6000)


def test_agg_kpis_noi_variance():
    kpis = [_kpi(10000, 4000)]
    result = agg_kpis(kpis)
    assert result["noi_variance"] is not None


def test_agg_kpis_empty_returns_all_none():
    result = agg_kpis([])
    assert result["actual_income"] is None
    assert result["actual_noi"] is None


def test_month_to_quarter():
    assert month_to_quarter(1) == 1
    assert month_to_quarter(3) == 1
    assert month_to_quarter(4) == 2
    assert month_to_quarter(12) == 4


def test_quarter_label():
    assert quarter_label(2025, 1) == "Q1 - 2025"
    assert quarter_label(2026, 4) == "Q4 - 2026"


def test_quarter_months():
    assert quarter_months(1) == {1, 2, 3}
    assert quarter_months(4) == {10, 11, 12}
