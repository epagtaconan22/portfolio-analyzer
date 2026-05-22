"""Parses a physical occupancy Excel report into OccupancyRow records."""

import openpyxl
from app.models import OccupancyRow

_PROPERTY_KEYWORDS = ["property", "prop", "name", "asset"]
_YEAR_KEYWORDS     = ["year", "yr"]
_MONTH_KEYWORDS    = ["month", "mo", "period"]
_OCCUPIED_KEYWORDS = ["occupied", "occ unit", "occ. unit", "units occ"]
_TOTAL_KEYWORDS    = ["total unit", "tot unit", "total units", "# units", "unit count"]

def parse_occupancy_report(file_path: str) -> list[OccupancyRow]:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    rows: list[OccupancyRow] = []

    for ws in wb.worksheets:
        all_cells = list(ws.iter_rows(values_only=True))
        header_idx, col_map = _find_header(all_cells)
        if header_idx is None:
            continue

        for row in all_cells[header_idx + 1:]:
            try:
                prop  = _get(row, col_map, "property")
                year  = int(_get(row, col_map, "year") or 0)
                month = int(_get(row, col_map, "month") or 0)
                occ   = int(_get(row, col_map, "occupied") or 0)
                total = int(_get(row, col_map, "total") or 0)
            except (TypeError, ValueError):
                continue
            if not prop or not year or not month or not total:
                continue
            rows.append(OccupancyRow(
                property_name=str(prop).strip(),
                year=year, month=month,
                occupied_units=occ, total_units=total,
            ))

    wb.close()
    return rows


def _find_header(all_cells) -> tuple:
    for idx, row in enumerate(all_cells[:20]):
        row_lower = [str(c).lower() if c else "" for c in row]
        col_map = {}
        for i, h in enumerate(row_lower):
            if any(k in h for k in _PROPERTY_KEYWORDS):
                col_map.setdefault("property", i)
            elif any(k in h for k in _YEAR_KEYWORDS):
                col_map.setdefault("year", i)
            elif any(k in h for k in _MONTH_KEYWORDS):
                col_map.setdefault("month", i)
            elif any(k in h for k in _OCCUPIED_KEYWORDS):
                col_map.setdefault("occupied", i)
            elif any(k in h for k in _TOTAL_KEYWORDS):
                col_map.setdefault("total", i)
        if all(k in col_map for k in ("property", "year", "month", "occupied", "total")):
            return idx, col_map
    return None, {}


def _get(row, col_map, key):
    idx = col_map.get(key)
    if idx is None or idx >= len(row):
        return None
    return row[idx]
