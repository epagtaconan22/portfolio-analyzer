# tests/test_financial_parser.py
import os

import pytest
from app.models import RawRow
from app.parser.financial import parse_financial_workbooks

def test_parses_actual_workbook(simple_actual_workbook):
    rows, source_index = parse_financial_workbooks(
        [simple_actual_workbook], {"test_actual.xlsx": "PM One"}
    )
    assert len(rows) > 0
    assert all(isinstance(r, RawRow) for r in rows)

def test_extracts_property_name_from_sheet(simple_actual_workbook):
    rows, _ = parse_financial_workbooks([simple_actual_workbook], {})
    props = {r.property_name for r in rows}
    assert any("sunrise" in p.lower() or "apts" in p.lower() for p in props)

def test_extracts_12_months(simple_actual_workbook):
    rows, _ = parse_financial_workbooks([simple_actual_workbook], {})
    gpr_rows = [r for r in rows if "gross potential" in r.account_name.lower()]
    months = {r.month for r in gpr_rows}
    assert months == set(range(1, 13))

def test_source_type_is_actual(simple_actual_workbook):
    rows, _ = parse_financial_workbooks([simple_actual_workbook], {})
    assert all(r.source_type == "Actual" for r in rows)

def test_parses_actual_budget_workbook(actual_budget_workbook):
    rows, _ = parse_financial_workbooks([actual_budget_workbook], {})
    source_types = {r.source_type for r in rows}
    assert "Actual" in source_types
    assert "Budget" in source_types

def test_source_index_documents_every_sheet(simple_actual_workbook):
    _, source_index = parse_financial_workbooks([simple_actual_workbook], {})
    assert len(source_index) >= 1
    assert all(hasattr(e, "source_sheet") for e in source_index)

def test_pm_name_assigned(simple_actual_workbook):
    rows, _ = parse_financial_workbooks(
        [simple_actual_workbook], {"test_actual.xlsx": "PM One"}
    )
    pm_rows = [r for r in rows if r.pm_name == "PM One"]
    assert len(pm_rows) > 0

def test_total_row_excluded(simple_actual_workbook):
    rows, _ = parse_financial_workbooks([simple_actual_workbook], {})
    account_names = [r.account_name.lower() for r in rows]
    assert not any(n.strip() in ("total", "total income", "total expenses", "net income")
                   for n in account_names)
