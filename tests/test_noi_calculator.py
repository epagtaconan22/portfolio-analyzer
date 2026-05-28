# tests/test_noi_calculator.py
import pytest
from app.models import MappedRow
from app.calculator.noi import calculate_noi, _safe_variance_pct

def _make_row(account_category, treatment, amount, year=2024, month=1,
              source_type="Actual", pm_name="PM Co", property_name="Prop A",
              account_name="Test Account"):
    return MappedRow(
        property_name=property_name, pm_name=pm_name,
        source_workbook="wb.xlsx", source_sheet="S1", source_type=source_type,
        source_row=1, account_code="", account_name=account_name,
        year=year, month=month, amount=amount, original_amount=amount,
        account_category=account_category, kpi_mapping=account_category,
        include_in_noi=account_category not in ("Excluded", "Review Needed"),
        include_in_eco_occ=account_category in ("Rental Income","Vacancy","Concessions","Bad Debt"),
        treatment=treatment,
    )

def test_noi_calculation():
    rows = [
        _make_row("Rental Income", "Income", 10000),
        _make_row("Other Income",  "Income", 500),
        _make_row("Vacancy",       "Contra-Income", 1000),
        _make_row("Bad Debt",      "Contra-Income", 200),
        _make_row("Operating Expense", "Expense", 4000),
    ]
    kpis = calculate_noi(rows)
    monthly = [k for k in kpis if k.month == 1 and k.property_name == "Prop A"]
    assert len(monthly) == 1
    k = monthly[0]
    # Income = 10000 + 500 - 1000 - 200 = 9300
    assert k.actual_income == pytest.approx(9300)
    # Expenses = 4000
    assert k.actual_expenses == pytest.approx(4000)
    # NOI = 9300 - 4000 = 5300
    assert k.actual_noi == pytest.approx(5300)

def test_noi_variance_against_budget():
    rows = [
        _make_row("Rental Income", "Income", 10000, source_type="Actual"),
        _make_row("Operating Expense", "Expense", 4000, source_type="Actual"),
        _make_row("Rental Income", "Income", 9000, source_type="Budget"),
        _make_row("Operating Expense", "Expense", 4500, source_type="Budget"),
    ]
    kpis = calculate_noi(rows)
    k = next(x for x in kpis if x.month == 1)
    # Actual NOI = 10000 - 4000 = 6000; Budget NOI = 9000 - 4500 = 4500
    assert k.actual_noi == pytest.approx(6000)
    assert k.budget_noi == pytest.approx(4500)
    assert k.noi_variance == pytest.approx(1500)
    assert k.noi_variance_pct == pytest.approx(1500 / 4500)

def test_noi_variance_pct_zero_budget():
    # When budget NOI = 0, variance % should be None (no divide by zero)
    assert _safe_variance_pct(500, 0) is None

def test_excluded_accounts_not_in_noi():
    rows = [
        _make_row("Rental Income", "Income", 10000),
        _make_row("Excluded", "Excluded", 50000),  # depreciation etc.
    ]
    kpis = calculate_noi(rows)
    k = next(x for x in kpis if x.month == 1)
    assert k.actual_noi == pytest.approx(10000)

def test_multi_property_separate_kpis():
    rows = [
        _make_row("Rental Income", "Income", 5000, property_name="Prop A"),
        _make_row("Rental Income", "Income", 8000, property_name="Prop B"),
    ]
    kpis = calculate_noi(rows)
    assert len(kpis) == 2
    props = {k.property_name for k in kpis}
    assert "Prop A" in props and "Prop B" in props

def test_expense_credit_reduces_expenses():
    """Negative expense rows (credits/reversals) must reduce expenses, not inflate them."""
    rows = [
        _make_row("Rental Income", "Income", 10000),
        _make_row("Operating Expense", "Expense", 4000),   # normal cost
        _make_row("Operating Expense", "Expense", -500),   # credit/reversal
    ]
    kpis = calculate_noi(rows)
    k = next(x for x in kpis if x.month == 1)
    assert k.actual_expenses == pytest.approx(3500)        # 4000 - 500
    assert k.actual_noi == pytest.approx(6500)             # 10000 - 3500

def test_top_two_drivers_identified():
    rows = [
        _make_row("Rental Income", "Income", 10000, source_type="Actual", account_name="Gross Potential Rent"),
        _make_row("Operating Expense", "Expense", 5000, source_type="Actual", account_name="Management Fee"),
        _make_row("Operating Expense", "Expense", 3000, source_type="Actual", account_name="Payroll"),
        _make_row("Rental Income", "Income", 9000, source_type="Budget", account_name="Gross Potential Rent"),
        _make_row("Operating Expense", "Expense", 3000, source_type="Budget", account_name="Management Fee"),
        _make_row("Operating Expense", "Expense", 4000, source_type="Budget", account_name="Payroll"),
    ]
    kpis = calculate_noi(rows)
    k = next(x for x in kpis if x.month == 1)
    # Management Fee: $-2000 unfavorable (largest); GPR: +$1000 and Payroll: +$1000 tied for second
    assert "Management Fee" in k.top_noi_driver_1
    assert k.top_noi_driver_2 != ""
