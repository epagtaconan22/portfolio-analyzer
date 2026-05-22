# tests/test_occupancy_parser.py
from app.models import OccupancyRow
from app.parser.occupancy import parse_occupancy_report

def test_parses_rows(occupancy_workbook):
    rows = parse_occupancy_report(occupancy_workbook)
    assert len(rows) == 3
    assert all(isinstance(r, OccupancyRow) for r in rows)

def test_property_names(occupancy_workbook):
    rows = parse_occupancy_report(occupancy_workbook)
    props = {r.property_name for r in rows}
    assert "Sunrise Apts" in props and "Oak Glen" in props

def test_physical_occ_pct(occupancy_workbook):
    rows = parse_occupancy_report(occupancy_workbook)
    sunrise_jan = next(r for r in rows if r.property_name == "Sunrise Apts" and r.month == 1)
    assert sunrise_jan.physical_occ_pct == 0.90

def test_flexible_column_headers(tmp_path):
    """Parser should handle variant column names."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Property Name", "Yr", "Mo", "Occ Units", "Tot Units"])
    ws.append(["Sunrise Apts", 2024, 3, 88, 100])
    path = tmp_path / "occ2.xlsx"
    wb.save(str(path))
    rows = parse_occupancy_report(str(path))
    assert len(rows) == 1
    assert rows[0].occupied_units == 88
