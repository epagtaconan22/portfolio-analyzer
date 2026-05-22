import pytest
from app import create_app

@pytest.fixture
def app():
    app = create_app()
    app.config["TESTING"] = True
    return app

@pytest.fixture
def client(app):
    return app.test_client()

import openpyxl

@pytest.fixture
def simple_actual_workbook(tmp_path):
    """A minimal actual income statement workbook for testing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Actual - Sunrise Apts"

    # Header row: account, Jan, Feb, Mar, ...Dec, Total
    ws.append(["Account", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total"])
    # Income rows
    ws.append(["Gross Potential Rent", 10000, 10000, 10200, 10200, 10200, 10200,
               10200, 10200, 10200, 10200, 10200, 10200, 122400])
    ws.append(["Vacancy Loss",         -500,  -600,  -400,  -400,  -400,  -400,
               -400,  -400,  -400,  -400,  -400,  -400,  -5100])
    ws.append(["Management Fee",       1200,  1200,  1200,  1200,  1200,  1200,
               1200,  1200,  1200,  1200,  1200,  1200,  14400])

    path = tmp_path / "test_actual.xlsx"
    wb.save(str(path))
    return str(path)

@pytest.fixture
def actual_budget_workbook(tmp_path):
    """A workbook with alternating Actual/Budget columns."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Actual vs Budget - Oak Glen"

    ws.append(["Account", "Jan Act", "Jan Bud", "Feb Act", "Feb Bud"])
    ws.append(["Gross Potential Rent", 10000, 10500, 10000, 10500])
    ws.append(["Vacancy Loss",         -500,  -300,  -600,  -300])

    path = tmp_path / "test_ab.xlsx"
    wb.save(str(path))
    return str(path)
