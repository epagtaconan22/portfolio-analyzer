# tests/test_sheet_inferrer.py
import pytest
from app.parser.sheet_inferrer import infer_sheet_type

def test_actual_by_sheet_name():
    assert infer_sheet_type("Actual Income Statement", [], []) == "Actual"

def test_budget_by_sheet_name():
    assert infer_sheet_type("Budget 2024", [], []) == "Budget"

def test_actual_budget_comparison_by_sheet_name():
    assert infer_sheet_type("Actual vs Budget", [], []) == "Actual+Budget"

def test_actual_by_header_keywords():
    headers = ["Account", "Jan Actual", "Feb Actual", "Mar Actual"]
    assert infer_sheet_type("Sheet1", headers, []) == "Actual"

def test_budget_by_header_keywords():
    headers = ["Account", "Jan Budget", "Feb Budget"]
    assert infer_sheet_type("Sheet1", headers, []) == "Budget"

def test_actual_budget_by_alternating_headers():
    headers = ["Account", "Jan Act", "Jan Bud", "Feb Act", "Feb Bud"]
    assert infer_sheet_type("Sheet1", headers, []) == "Actual+Budget"

def test_unknown_sheet():
    assert infer_sheet_type("Notes", [], []) == "Unknown"

def test_title_row_actual():
    # First non-empty cell in sheet often contains "Income Statement - Actual"
    title_rows = [["Income Statement - Actual YTD 2024"]]
    assert infer_sheet_type("Sheet1", [], title_rows) == "Actual"
