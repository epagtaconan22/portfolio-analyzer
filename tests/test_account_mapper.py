# tests/test_account_mapper.py
from app.mapper.account_mapper import map_account_name, map_rows
from app.models import RawRow, MappedRow

def _make_raw(account_name, amount=1000.0, source_type="Actual"):
    return RawRow(
        property_name="Test Property", pm_name="PM Co", source_workbook="wb.xlsx",
        source_sheet="Sheet1", source_type=source_type, source_row=1,
        account_code="", account_name=account_name, year=2024, month=1,
        amount=amount, original_amount=amount,
    )

def test_maps_rental_income():
    cat, treatment, in_noi, in_eco = map_account_name("Gross Potential Rent")
    assert cat == "Rental Income"
    assert treatment == "Income"
    assert in_noi is True
    assert in_eco is True

def test_maps_vacancy():
    cat, treatment, in_noi, in_eco = map_account_name("Vacancy Loss")
    assert cat == "Vacancy"
    assert treatment == "Contra-Income"
    assert in_noi is True
    assert in_eco is True

def test_maps_bad_debt():
    cat, treatment, in_noi, in_eco = map_account_name("Bad Debt Expense")
    assert cat == "Bad Debt"
    assert treatment == "Contra-Income"
    assert in_noi is True
    assert in_eco is True

def test_maps_concession():
    cat, treatment, in_noi, in_eco = map_account_name("Rent Concessions")
    assert cat == "Concessions"
    assert treatment == "Contra-Income"
    assert in_noi is True
    assert in_eco is True

def test_maps_operating_expense():
    cat, treatment, in_noi, _ = map_account_name("Management Fee")
    assert cat == "Operating Expense"
    assert treatment == "Expense"
    assert in_noi is True

def test_maps_excluded():
    cat, treatment, in_noi, _ = map_account_name("Depreciation Expense")
    assert cat == "Excluded"
    assert in_noi is False

def test_unmatched_is_review_needed():
    cat, _, _, _ = map_account_name("XYZ Special Item")
    assert cat == "Review Needed"

def test_case_insensitive():
    cat, _, _, _ = map_account_name("GROSS POTENTIAL RENT")
    assert cat == "Rental Income"

def test_custom_mapping_overrides_default():
    custom = {"custom special item": ("Other Income", "Income", True, False)}
    cat, _, _, _ = map_account_name("Custom Special Item", custom_mapping=custom)
    assert cat == "Other Income"

def test_map_rows_returns_mapped_rows():
    rows = [_make_raw("Gross Potential Rent"), _make_raw("Vacancy Loss")]
    mapped, entries = map_rows(rows)
    assert len(mapped) == 2
    assert all(isinstance(r, MappedRow) for r in mapped)
    assert mapped[0].account_category == "Rental Income"
    assert mapped[1].account_category == "Vacancy"
    assert len(entries) == 2

def test_map_rows_kpi_mapping_field():
    rows = [_make_raw("Gross Potential Rent")]
    mapped, _ = map_rows(rows)
    assert mapped[0].kpi_mapping == "GPR / Rental Income"

def test_map_rows_deduplication():
    # Same account in two months → one MappingEntry
    r1 = RawRow(
        property_name="P", pm_name="PM", source_workbook="w.xlsx",
        source_sheet="S1", source_type="Actual", source_row=1,
        account_code="5000", account_name="Gross Potential Rent",
        year=2024, month=1, amount=10000, original_amount=10000,
    )
    r2 = RawRow(
        property_name="P", pm_name="PM", source_workbook="w.xlsx",
        source_sheet="S1", source_type="Actual", source_row=2,
        account_code="5000", account_name="Gross Potential Rent",
        year=2024, month=2, amount=10000, original_amount=10000,
    )
    mapped, entries = map_rows([r1, r2])
    assert len(mapped) == 2
    assert len(entries) == 1
    assert entries[0].assigned_category == "Rental Income"
    assert entries[0].include_in_noi is True
    assert entries[0].include_in_eco_occ is True

def test_misc_expense_maps_to_operating_expense():
    cat, treatment, _, _ = map_account_name("Miscellaneous Expense")
    assert cat == "Operating Expense"
    assert treatment == "Expense"

def test_security_deposit_maps_to_other_income():
    cat, treatment, _, _ = map_account_name("Security Deposit Forfeiture")
    assert cat == "Other Income"
    assert treatment == "Income"
