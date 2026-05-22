"""Parses financial statement Excel workbooks into RawRow records."""

import os
import re
from typing import Optional
import openpyxl
from app.models import RawRow, SourceIndexEntry
from app.parser.sheet_inferrer import infer_sheet_type

_MONTH_PATTERNS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "january":1,"february":2,"march":3,"april":4,"june":6,"july":7,
    "august":8,"september":9,"october":10,"november":11,"december":12,
}

_SKIP_ACCOUNT_PATTERNS = {
    "total", "subtotal", "net income", "net loss", "total income",
    "total revenue", "total expenses", "total cost", "total operating",
    "grand total", "net operating",
}

def parse_financial_workbooks(
    file_paths: list[str],
    pm_name_map: dict[str, str],   # filename (basename) → PM company name
) -> tuple[list[RawRow], list[SourceIndexEntry]]:
    all_rows: list[RawRow] = []
    source_index: list[SourceIndexEntry] = []

    for path in file_paths:
        fname = os.path.basename(path)
        pm_name = pm_name_map.get(fname, _infer_pm_from_filename(fname))
        wb = openpyxl.load_workbook(path, data_only=True)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows, entry = _parse_sheet(ws, sheet_name, fname, pm_name)
            all_rows.extend(rows)
            source_index.append(entry)

    return all_rows, source_index


def _parse_sheet(
    ws, sheet_name: str, source_workbook: str, pm_name: str
) -> tuple[list[RawRow], SourceIndexEntry]:
    all_cells = list(ws.iter_rows(values_only=True))
    if not all_cells:
        return [], SourceIndexEntry(
            source_workbook=source_workbook, source_sheet=sheet_name,
            property_name="", pm_name=pm_name, year=0, source_type="Unknown",
            processed=False, rows_extracted=0, reason_if_excluded="Empty sheet",
        )

    title_rows = [list(r) for r in all_cells[:5]]
    header_row_idx, header_row = _find_header_row(all_cells)
    if header_row_idx is None:
        return [], SourceIndexEntry(
            source_workbook=source_workbook, source_sheet=sheet_name,
            property_name="", pm_name=pm_name, year=0, source_type="Unknown",
            processed=False, rows_extracted=0, reason_if_excluded="No month header row found",
        )

    header_strs = [str(c) if c is not None else "" for c in header_row]
    source_type = infer_sheet_type(sheet_name, header_strs, title_rows)
    property_name = _infer_property_name(sheet_name, title_rows, source_type)
    year = _infer_year(sheet_name, title_rows, header_strs)

    # Map column index → (month, stream) where stream = "Actual" or "Budget"
    col_map = _map_month_columns(header_strs, source_type)
    if not col_map:
        return [], SourceIndexEntry(
            source_workbook=source_workbook, source_sheet=sheet_name,
            property_name=property_name, pm_name=pm_name, year=year,
            source_type=source_type, processed=False, rows_extracted=0,
            reason_if_excluded="No month columns identified",
        )

    rows: list[RawRow] = []
    for row_idx, row in enumerate(all_cells[header_row_idx + 1:], start=header_row_idx + 2):
        account_code, account_name = _extract_account(row)
        if not account_name or _is_skip_row(account_name):
            continue

        for col_idx, (month, stream) in col_map.items():
            if col_idx >= len(row):
                continue
            raw_val = row[col_idx]
            if raw_val is None:
                continue
            try:
                amount = float(raw_val)
            except (ValueError, TypeError):
                continue

            rows.append(RawRow(
                property_name=property_name,
                pm_name=pm_name,
                source_workbook=source_workbook,
                source_sheet=sheet_name,
                source_type=stream,
                source_row=row_idx,
                account_code=account_code,
                account_name=account_name,
                year=year,
                month=month,
                amount=amount,
                original_amount=amount,
            ))

    return rows, SourceIndexEntry(
        source_workbook=source_workbook, source_sheet=sheet_name,
        property_name=property_name, pm_name=pm_name, year=year,
        source_type=source_type, processed=True, rows_extracted=len(rows),
    )


def _find_header_row(all_cells) -> tuple[Optional[int], Optional[tuple]]:
    """Find the row index that contains month abbreviations (Jan/Feb/etc.)."""
    for idx, row in enumerate(all_cells[:20]):
        row_strs = [str(c).lower() if c else "" for c in row]
        month_hits = sum(1 for s in row_strs if any(m in s for m in _MONTH_PATTERNS))
        if month_hits >= 3:
            return idx, row
    return None, None


def _map_month_columns(header_strs: list[str], source_type: str) -> dict[int, tuple[int, str]]:
    """Returns {col_index: (month_int, 'Actual'|'Budget')}."""
    col_map: dict[int, tuple[int, str]] = {}
    for idx, h in enumerate(header_strs):
        h_lower = h.lower().strip()
        month = _extract_month(h_lower)
        if month is None:
            continue
        if source_type == "Actual+Budget":
            stream = "Budget" if ("bud" in h_lower or "budget" in h_lower) else "Actual"
        elif source_type == "Budget":
            stream = "Budget"
        else:
            stream = "Actual"
        col_map[idx] = (month, stream)
    return col_map


def _extract_month(s: str) -> Optional[int]:
    for abbr, num in _MONTH_PATTERNS.items():
        if abbr in s:
            return num
    # Try MM/YYYY or MM-YYYY pattern
    m = re.search(r'\b(0?[1-9]|1[0-2])[/\-](\d{4})\b', s)
    if m:
        return int(m.group(1))
    return None


def _extract_account(row) -> tuple[str, str]:
    """Returns (account_code, account_name) from first 1-3 cells of a data row."""
    code = ""
    name = ""
    for cell in row[:3]:
        if cell is None:
            continue
        val = str(cell).strip()
        if not val:
            continue
        # If first non-empty cell looks like an account code (e.g. "5100"), save as code
        if not name and re.match(r'^\d{3,6}$', val):
            code = val
        elif not name:
            name = val
        elif not code and re.match(r'^\d{3,6}$', val):
            code = val
        else:
            break
    return code, name


def _is_skip_row(account_name: str) -> bool:
    return account_name.lower().strip() in _SKIP_ACCOUNT_PATTERNS


def _infer_property_name(sheet_name: str, title_rows: list, source_type: str) -> str:
    """Extract property name from sheet name (after removing source type keywords)."""
    name = sheet_name
    for kw in ["actual vs budget", "actual/budget", "actual", "budget", "-", "–"]:
        name = re.sub(re.escape(kw), " ", name, flags=re.IGNORECASE)
    name = name.strip(" -–|")
    if name:
        return name
    # Fall back to first non-empty title row cell
    for row in title_rows[:3]:
        for cell in row[:3]:
            if cell:
                return str(cell).strip()
    return "Unknown Property"


def _infer_year(sheet_name: str, title_rows: list, header_strs: list[str]) -> int:
    """Extract year from sheet name, title, or column headers."""
    for src in [sheet_name] + [str(c) for row in title_rows[:3] for c in row if c] + header_strs:
        m = re.search(r'\b(20\d{2})\b', str(src))
        if m:
            return int(m.group(1))
    return 0


def _infer_pm_from_filename(filename: str) -> str:
    return os.path.splitext(filename)[0].replace("_", " ").replace("-", " ").strip()
