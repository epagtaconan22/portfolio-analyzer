"""Parses financial statement Excel workbooks into RawRow records."""

import datetime as _dt
import itertools
import os
import re
from typing import Optional
import openpyxl
from app.models import RawRow, SourceIndexEntry
from app.parser.sheet_inferrer import (
    infer_sheet_type,
    _ACTUAL_BUDGET_PATTERNS,
    _ACTUAL_PATTERNS,
    _BUDGET_PATTERNS,
)

# Short 3-letter abbreviations are sufficient — substring matching means "jan" already
# matches "january", "jun" matches "june", etc.  Long-form duplicates are not needed.
_MONTH_PATTERNS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Compiled once for use in _find_header_row — avoids per-cell linear key scan.
_MONTH_RE = re.compile(r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b')

# Matches MM/YYYY, MM-YYYY, YYYY/MM, YYYY-MM (numeric date column headers).
_NUMERIC_MONTH_RE = re.compile(
    r'\b(?:0?[1-9]|1[0-2])[/\-](?:20\d{2})\b'   # MM/YYYY or MM-YYYY
    r'|'
    r'\b(?:20\d{2})[/\-](?:0?[1-9]|1[0-2])\b'   # YYYY/MM or YYYY-MM
)

# Keywords to strip from sheet names when inferring a property name, built from the
# same source as sheet_inferrer so the two modules stay in sync.
_SHEET_NAME_STRIP_PATTERNS = _ACTUAL_BUDGET_PATTERNS + _ACTUAL_PATTERNS + _BUDGET_PATTERNS
_SHEET_NAME_STRIP_RE = re.compile(
    "|".join(re.escape(p) for p in sorted(_SHEET_NAME_STRIP_PATTERNS, key=len, reverse=True)),
    flags=re.IGNORECASE,
)

# Account code: starts with a digit, followed by 2+ digits or common separators.
# Handles short codes ("500"), long Yardi codes ("5150000"), and dash/dot formats ("5000-01").
_ACCOUNT_CODE_RE = re.compile(r'^\d[\d\-\.:/]{2,}$')

_SKIP_ACCOUNT_PATTERNS = {
    "total", "subtotal", "net income", "net loss", "total income",
    "total revenue", "total expenses", "total cost", "total operating",
    "grand total", "net operating",
}

_YEAR_RE = re.compile(r'\b(20\d{2})\b')


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


def _make_index_entry(
    source_workbook: str,
    sheet_name: str,
    pm_name: str,
    *,
    property_name: str = "",
    year: int = 0,
    source_type: str = "Unknown",
    processed: bool = False,
    rows_extracted: int = 0,
    reason_if_excluded: str = "",
) -> SourceIndexEntry:
    """Single construction point for SourceIndexEntry to avoid copy-paste."""
    return SourceIndexEntry(
        source_workbook=source_workbook,
        source_sheet=sheet_name,
        property_name=property_name,
        pm_name=pm_name,
        year=year,
        source_type=source_type,
        processed=processed,
        rows_extracted=rows_extracted,
        reason_if_excluded=reason_if_excluded,
    )


def _parse_sheet(
    ws, sheet_name: str, source_workbook: str, pm_name: str
) -> tuple[list[RawRow], SourceIndexEntry]:
    all_cells = list(ws.iter_rows(values_only=True))
    if not all_cells:
        return [], _make_index_entry(
            source_workbook, sheet_name, pm_name,
            reason_if_excluded="Empty sheet",
        )

    title_rows = [list(r) for r in all_cells[:5]]
    header_row_idx, header_row = _find_header_row(all_cells)
    if header_row_idx is None:
        return [], _make_index_entry(
            source_workbook, sheet_name, pm_name,
            reason_if_excluded="No month header row found",
        )

    header_strs = [_cell_to_str(c) for c in header_row]
    source_type = infer_sheet_type(sheet_name, header_strs, title_rows)
    property_name = _infer_property_name(sheet_name, title_rows)
    year = _infer_year(sheet_name, title_rows, header_strs)

    col_map = _map_month_columns(header_strs, source_type)
    if not col_map:
        return [], _make_index_entry(
            source_workbook, sheet_name, pm_name,
            property_name=property_name, year=year, source_type=source_type,
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

    return rows, _make_index_entry(
        source_workbook, sheet_name, pm_name,
        property_name=property_name, year=year, source_type=source_type,
        processed=True, rows_extracted=len(rows),
    )


def _find_header_row(all_cells) -> tuple[Optional[int], Optional[tuple]]:
    """Return the first row (within the first 30) that contains 3+ month indicators.

    Handles text abbreviations ("Jan", "February"), numeric date headers ("1/2024",
    "2024-01"), and Excel date values returned by openpyxl as datetime objects.
    """
    for idx, row in enumerate(all_cells[:30]):
        hits = 0
        for c in row:
            if c is None:
                continue
            if isinstance(c, (_dt.date, _dt.datetime)):
                hits += 1
                continue
            s = str(c).lower()
            if _MONTH_RE.search(s) or _NUMERIC_MONTH_RE.search(s):
                hits += 1
        if hits >= 3:
            return idx, row
    return None, None


def _map_month_columns(header_strs: list[str], source_type: str) -> dict[int, tuple[int, str]]:
    """Return {col_index: (month_int, stream)} where stream is 'Actual' or 'Budget'.

    For Actual+Budget sheets the stream is determined per-column by whether the header
    contains a budget keyword; for pure-Budget sheets every month column is 'Budget';
    for all other types every month column is 'Actual'.
    """
    # Resolve the stream-assignment strategy once, outside the column loop.
    if source_type == "Actual+Budget":
        def _stream(h_lower: str) -> str:
            return "Budget" if ("bud" in h_lower) else "Actual"
    elif source_type == "Budget":
        def _stream(h_lower: str) -> str:  # noqa: F811
            return "Budget"
    else:
        def _stream(h_lower: str) -> str:  # noqa: F811
            return "Actual"

    col_map: dict[int, tuple[int, str]] = {}
    for idx, h in enumerate(header_strs):
        h_lower = h.lower().strip()
        month = _extract_month(h_lower)
        if month is not None:
            col_map[idx] = (month, _stream(h_lower))
    return col_map


def _cell_to_str(c) -> str:
    """Convert a cell value to a string suitable for month/year detection.

    datetime objects are formatted as "Jan 2024" so _MONTH_RE and _YEAR_RE
    can match them without special-casing every downstream function.
    """
    if c is None:
        return ""
    if isinstance(c, (_dt.date, _dt.datetime)):
        return c.strftime("%b %Y")   # e.g. "Jan 2024"
    return str(c)


def _extract_month(s: str) -> Optional[int]:
    m = _MONTH_RE.search(s.lower())
    if m:
        return _MONTH_PATTERNS[m.group()]
    # MM/YYYY or MM-YYYY
    m2 = re.search(r'\b(0?[1-9]|1[0-2])[/\-](20\d{2})\b', s)
    if m2:
        return int(m2.group(1))
    # YYYY/MM or YYYY-MM
    m3 = re.search(r'\b(20\d{2})[/\-](0?[1-9]|1[0-2])\b', s)
    if m3:
        return int(m3.group(2))
    return None


def _extract_account(row) -> tuple[str, str]:
    """Return (account_code, account_name) from the first 1-4 cells of a data row."""
    code = ""
    name = ""
    for cell in row[:4]:
        if cell is None:
            continue
        val = str(cell).strip()
        if not val:
            continue
        if not name and _ACCOUNT_CODE_RE.match(val):
            code = val
        elif not name:
            name = val
        elif not code and _ACCOUNT_CODE_RE.match(val):
            code = val
        else:
            break
    return code, name


def _is_skip_row(account_name: str) -> bool:
    lower = account_name.lower().strip()
    if lower in _SKIP_ACCOUNT_PATTERNS:
        return True
    # Skip Yardi-style subtotal rows whose names begin with "Total " or "Net "
    # (e.g. "Total Rent Revenue", "Net Rental Revenue", "NET OPERATING INCOME (Loss)").
    # Individual line-item accounts almost never start with these words.
    if lower.startswith(("total ", "net ")):
        return True
    return False


# Matches Yardi title-row pattern: "Property Name (short_code)" where the
# parenthetical is 2–15 alphanumeric/hyphen/underscore characters.
_YARDI_TITLE_RE = re.compile(r'^(.+?)\s*\([A-Za-z0-9_\-]{2,15}\)\s*$')


def _infer_property_name(sheet_name: str, title_rows: list) -> str:
    """Extract property name, preferring the Yardi-style title row when present.

    Yardi exports place a "Property Name (code)" string in the first cell of row 1
    for every sheet.  Matching that pattern produces a clean, consistent name
    regardless of whether the sheet is a Statement or Budget sheet.
    """
    # 1. Check the first two title rows for the Yardi "Name (code)" pattern.
    for row in title_rows[:2]:
        if not row:
            continue
        cell = row[0]
        if not cell:
            continue
        val = str(cell).strip()
        m = _YARDI_TITLE_RE.match(val)
        if m:
            return m.group(1).strip()

    # 2. Fall back to sheet-name stripping (covers "Actual - Sunrise Apts" etc.).
    name = _SHEET_NAME_STRIP_RE.sub(" ", sheet_name).strip(" -–|").strip()
    if name:
        return name

    # 3. Last resort: first non-empty cell in the title area.
    for row in title_rows[:3]:
        for cell in row[:3]:
            if cell:
                return str(cell).strip()
    return "Unknown Property"


def _infer_year(sheet_name: str, title_rows: list, header_strs: list[str]) -> int:
    """Extract the first four-digit year (20xx) from sheet name, title rows, or headers."""
    title_cells = (str(c) for row in title_rows[:3] for c in row if c)
    for src in itertools.chain([sheet_name], title_cells, header_strs):
        m = _YEAR_RE.search(src)
        if m:
            return int(m.group(1))
    return 0


def _infer_pm_from_filename(filename: str) -> str:
    return os.path.splitext(filename)[0].replace("_", " ").replace("-", " ").strip()
