import pytest
from app.models import MappedRow, PropertyPeriodKPIs
from app.calculator.economic_occ import enrich_eco_occ

def _mapped(cat, treatment, amount, source_type="Actual", property_name="Prop A"):
    return MappedRow(
        property_name=property_name, pm_name="PM", source_workbook="wb.xlsx",
        source_sheet="S1", source_type=source_type, source_row=1,
        account_code="", account_name=cat, year=2024, month=1,
        amount=amount, original_amount=amount,
        account_category=cat, kpi_mapping=cat,
        include_in_noi=True,
        include_in_eco_occ=cat in ("Rental Income","Vacancy","Concessions","Bad Debt"),
        treatment=treatment,
    )

def _base_kpi(property_name="Prop A"):
    return PropertyPeriodKPIs(property_name=property_name, pm_name="PM",
                               year=2024, month=1, period="Jan")

def test_eco_occ_basic():
    rows = [
        _mapped("Rental Income", "Income",        10000),
        _mapped("Vacancy",       "Contra-Income",  500),
        _mapped("Concessions",   "Contra-Income",  200),
        _mapped("Bad Debt",      "Contra-Income",  300),
    ]
    kpis = [_base_kpi()]
    result = enrich_eco_occ(rows, kpis)
    k = result[0]
    assert k.gpr == pytest.approx(10000)
    assert k.vacancy == pytest.approx(500)
    assert k.concessions == pytest.approx(200)
    assert k.bad_debt == pytest.approx(300)
    assert k.net_collectible == pytest.approx(9000)
    assert k.eco_occ_pct == pytest.approx(0.90)

def test_eco_occ_100_percent_no_losses():
    rows = [_mapped("Rental Income", "Income", 10000)]
    kpis = [_base_kpi()]
    result = enrich_eco_occ(rows, kpis)
    assert result[0].eco_occ_pct == pytest.approx(1.0)

def test_eco_occ_flagged_if_over_100():
    # GPR = 10000, but vacancy is negative in source (sign error) → net_collectible > GPR
    rows = [
        _mapped("Rental Income", "Income",  10000),
        _mapped("Vacancy",       "Contra-Income", -500),  # negative vacancy = unusual
    ]
    kpis = [_base_kpi()]
    result = enrich_eco_occ(rows, kpis)
    # eco_occ > 1.0 should still be calculated (flagged separately in Quality_Checks)
    assert result[0].eco_occ_pct is not None
    assert result[0].eco_occ_pct > 1.0

def test_eco_occ_none_if_no_gpr():
    rows = [_mapped("Vacancy", "Contra-Income", 500)]
    kpis = [_base_kpi()]
    result = enrich_eco_occ(rows, kpis)
    assert result[0].eco_occ_pct is None

def test_budget_eco_occ():
    rows = [
        _mapped("Rental Income", "Income",       10000, source_type="Budget"),
        _mapped("Vacancy",       "Contra-Income",  800, source_type="Budget"),
    ]
    kpis = [_base_kpi()]
    result = enrich_eco_occ(rows, kpis)
    assert result[0].budget_eco_occ_pct == pytest.approx(0.92)
