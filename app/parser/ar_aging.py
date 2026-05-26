"""Parses Yardi AR Aging export files (Tenant Rent and Subsidy) into ARAgingRow records."""

import os
import re
from typing import Optional
import openpyxl

from app.models import ARAgingRow

# Map lowercase type strings from filename to canonical values
_TYPE_NORMALIZE: dict[str, str] = {
    "tenant rent":        "Tenant Rent",
    "tenant receivable":  "Tenant Rent",
    "subsidy":            "Subsidy",
    "subsidy receivable": "Subsidy",
}


def parse_ar_aging_reports(file_paths: list[str]) -> list[ARAgingRow]:
    """Parse one or more Yardi AR Aging export files. Returns combined list of ARAgingRow."""
    results: list[ARAgingRow] = []
    for path in file_paths:
        results.extend(_parse_one(path))
    return results


def _parse_one(path: str) -> list[ARAgingRow]:
    fname = os.path.basename(path)
    stem  = os.path.splitext(fname)[0]

    pm_name, receivable_type, year, month = _parse_filename(stem)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Report1"] if "Report1" in wb.sheetnames else wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    # Fallback: read period from sheet row 2 if filename didn't match
    if year is None or month is None:
        year, month = _parse_period_from_sheet(all_rows)
    if receivable_type is None:
        receivable_type = _infer_type_from_stem(stem)

    rows: list[ARAgingRow] = []
    # Data rows start at index 5 (0-based) — after 5 header rows
    for raw_row in all_rows[5:]:
        col_a = raw_row[0]
        if col_a is None:
            break
        raw_name = str(col_a).strip()
        if not raw_name or raw_name.startswith("Grand Total"):
            break

        # Strip property code suffix "(code)" from property name
        property_name = re.sub(r'\s*\([^)]+\)\s*$', '', raw_name).strip()
        if not property_name:
            continue

        rows.append(ARAgingRow(
            property_name=property_name,
            pm_name=pm_name or "",
            source_file=fname,
            receivable_type=receivable_type or "Unknown",
            year=year or 0,
            month=month or 0,
            charge_amount=_to_float(raw_row[1]),
            current_owed=_to_float(raw_row[2]),
            owed_0_30=_to_float(raw_row[3]),
            owed_31_60=_to_float(raw_row[4]),
            owed_61_90=_to_float(raw_row[5]),
            owed_over_90=_to_float(raw_row[6]),
            prepayments=_to_float(raw_row[7]),
        ))

    return rows


def _parse_filename(stem: str) -> tuple[Optional[str], Optional[str], Optional[int], Optional[int]]:
    """
    Parse stem like "Solari_AR Aging_Tenant Rent_03_2024".
    Returns (pm_name, receivable_type, year, month).
    Returns (None, None, None, None) if the pattern doesn't match.
    """
    parts = stem.split("_")
    if len(parts) < 5:
        return None, None, None, None

    try:
        month = int(parts[-2])
        year  = int(parts[-1])
    except ValueError:
        return None, None, None, None

    if not (1 <= month <= 12 and 2000 <= year <= 2100):
        return None, None, None, None

    pm_name  = parts[0]
    type_raw = " ".join(parts[2:-2]).lower().strip()
    receivable_type = _TYPE_NORMALIZE.get(type_raw)

    return pm_name, receivable_type, year, month


def _parse_period_from_sheet(all_rows: list) -> tuple[Optional[int], Optional[int]]:
    """Fallback: parse period from row index 2 — 'Post To(MM/YY): 03/2024'."""
    if len(all_rows) < 3 or all_rows[2][0] is None:
        return None, None
    m = re.search(r'(\d{1,2})/(\d{4})', str(all_rows[2][0]))
    if m:
        return int(m.group(2)), int(m.group(1))   # (year, month)
    return None, None


def _infer_type_from_stem(stem: str) -> str:
    """Infer receivable type from filename stem keywords when pattern doesn't match."""
    stem_lower = stem.lower()
    if "subsidy" in stem_lower:
        return "Subsidy"
    return "Tenant Rent"


def _to_float(val) -> float:
    """Convert a cell value to float; treat None as 0.0."""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
