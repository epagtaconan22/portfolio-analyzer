"""Builds the 3-tab presentation workbook: Dashboard, Property Analysis, Property Monthly KPIs."""

import os
from typing import Optional
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from app.models import PropertyPeriodKPIs
from app.exporter.styles import (
    style_header_row, add_kpi_comment, apply_variance_fill,
    CURRENCY_FMT, PCT_FMT, VAR_PCT_FMT, COMMA_FMT, BOLD_FONT,
    SUBHDR_FILL, SUBHEADER_FONT,
)
from config import ECO_OCC_TARGET, KPI_FORMULAS


# ─── Quarter helpers ─────────────────────────────────────────────────────────

def _month_to_quarter(month: int) -> int:
    """Return quarter number (1–4) for a given month (1–12)."""
    return (month - 1) // 3 + 1


def _quarter_label(year: int, quarter: int) -> str:
    """Return display label, e.g. 'Q1 - 2025'."""
    return f"Q{quarter} - {year}"


def _get_sorted_quarters(kpis) -> list[tuple[int, int]]:
    """Return sorted list of (year, quarter) tuples present in the KPI list."""
    quarters: set[tuple[int, int]] = set()
    for k in kpis:
        quarters.add((k.year, _month_to_quarter(k.month)))
    return sorted(quarters)


def _kpis_for_quarter(kpis, year: int, quarter: int) -> list:
    """Return KPI records belonging to the given year and quarter."""
    months = {(quarter - 1) * 3 + 1, (quarter - 1) * 3 + 2, (quarter - 1) * 3 + 3}
    return [k for k in kpis if k.year == year and k.month in months]


# ─── Dashboard KPI row definitions ───────────────────────────────────────────
# Each entry: (display_label, _aggregate()-dict-key, number_format)
# None entries produce a blank separator row in the transposed table.

_DASHBOARD_KPI_ROWS = [
    ("Actual Income",       "actual_income",        CURRENCY_FMT),
    ("Budget Income",       "budget_income",        CURRENCY_FMT),
    ("Income Variance",     "income_variance",      CURRENCY_FMT),
    ("Income Variance %",   "income_variance_pct",  PCT_FMT),
    None,
    ("Actual Expenses",     "actual_expenses",      CURRENCY_FMT),
    ("Budget Expenses",     "budget_expenses",      CURRENCY_FMT),
    ("Expense Variance",    "expense_variance",     CURRENCY_FMT),
    ("Expense Variance %",  "expense_variance_pct", PCT_FMT),
    None,
    ("Actual NOI",          "actual_noi",           CURRENCY_FMT),
    ("Budget NOI",          "budget_noi",           CURRENCY_FMT),
    ("NOI Variance",        "noi_variance",         CURRENCY_FMT),
    ("NOI Variance %",      "noi_variance_pct",     PCT_FMT),
    None,
    ("GPR",                 "gpr",                  CURRENCY_FMT),
    ("Vacancy",             "vacancy",              CURRENCY_FMT),
    ("Concessions",         "concessions",          CURRENCY_FMT),
    ("Bad Debt",            "bad_debt",             CURRENCY_FMT),
    ("Net Collectible",     "net_collectible",      CURRENCY_FMT),
    ("Eco Occ %",           "eco_occ_pct",          PCT_FMT),
    ("Budget Eco Occ %",    "budget_eco_occ_pct",   PCT_FMT),
    ("Eco Occ Variance",    "eco_occ_variance",     PCT_FMT),
    None,
    ("Physical Occ %",      "physical_occ_pct",     PCT_FMT),
    ("Leakage Gap",         "leakage_gap",          PCT_FMT),
    None,
    ("Income/Unit",         "income_per_unit",      CURRENCY_FMT),
    ("Expense/Unit",        "expense_per_unit",     CURRENCY_FMT),
    ("NOI/Unit",            "noi_per_unit",         CURRENCY_FMT),
]


# ─── Public entry point ───────────────────────────────────────────────────────

def build_main_workbook(
    kpis: list[PropertyPeriodKPIs],
    portfolio_name: str,
    output_path: str,
    eco_occ_target: float = ECO_OCC_TARGET,
) -> str:
    """Builds main workbook at output_path. Returns path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_dashboard(wb, kpis, portfolio_name, eco_occ_target)
    _build_property_analysis(wb, kpis, portfolio_name, eco_occ_target)
    _build_monthly_kpis(wb, kpis)

    wb.save(output_path)
    return output_path


# ─── Dashboard ────────────────────────────────────────────────────────────────

def _build_dashboard(wb, kpis, portfolio_name, eco_occ_target):
    ws = wb.create_sheet("Dashboard")

    props = {k.property_name for k in kpis if not k.is_carveout}
    num_props = len(props)

    # ── Quarter-period aggregates ────────────────────────────────────────────
    quarters = _get_sorted_quarters(kpis)
    period_labels = [_quarter_label(yr, q) for (yr, q) in quarters]
    period_aggs: dict[str, dict] = {}
    for (yr, q) in quarters:
        q_kpis = [k for k in _kpis_for_quarter(kpis, yr, q) if not k.is_carveout]
        period_aggs[_quarter_label(yr, q)] = _aggregate(q_kpis)

    # ── Title ────────────────────────────────────────────────────────────────
    row = 1
    ws.cell(row, 1, f"{portfolio_name} — Portfolio Summary ({num_props} Properties)")
    ws.cell(row, 1).font = BOLD_FONT
    row += 2

    # ── Transposed KPI table: KPI label | Q1-2024 | Q2-2024 | ... ───────────
    num_period_cols = len(period_labels)
    total_cols = 1 + num_period_cols

    # Header row
    ws.cell(row, 1, "KPI")
    for col_idx, lbl in enumerate(period_labels, 2):
        cell = ws.cell(row, col_idx, lbl)
        cell.alignment = Alignment(horizontal="center")
    style_header_row(ws, row, total_cols)
    row += 1

    # KPI data rows
    for entry in _DASHBOARD_KPI_ROWS:
        if entry is None:
            row += 1
            continue
        label, key, fmt = entry
        cell = ws.cell(row, 1, label)
        add_kpi_comment(cell, label)
        for col_idx, period_lbl in enumerate(period_labels, 2):
            val = period_aggs[period_lbl].get(key)
            c = ws.cell(row, col_idx, val)
            if fmt:
                c.number_format = fmt
        row += 1

    row += 2

    # ── NOI Trend Commentary ─────────────────────────────────────────────────
    ws.cell(row, 1, "NOI Trend Commentary").font = BOLD_FONT
    row += 1
    commentary = _generate_noi_commentary(kpis)
    ws.cell(row, 1, commentary)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 40
    # Set period columns to a comfortable width
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
    row += 3

    # ── Top 5 Positive NOI Variance ──────────────────────────────────────────
    ws.cell(row, 1, f"Top Positive NOI Variances ({num_props} Properties Analyzed)").font = BOLD_FONT
    row += 1
    row = _write_top_noi_table(ws, kpis, row, top_n=5, favorable=True)
    row += 2

    # ── Top 5 Negative NOI Variance ──────────────────────────────────────────
    ws.cell(row, 1, f"Top Negative NOI Variances ({num_props} Properties Analyzed)").font = BOLD_FONT
    row += 1
    row = _write_top_noi_table(ws, kpis, row, top_n=5, favorable=False)
    row += 2

    # ── Properties Below Eco Occ Target ─────────────────────────────────────
    ws.cell(row, 1, f"Properties Below Economic Occupancy Target ({eco_occ_target:.0%})").font = BOLD_FONT
    row += 1
    _write_below_target_table(ws, kpis, row, eco_occ_target)

    # No frozen panes — different sections have different column layouts


# ─── Property Analysis ────────────────────────────────────────────────────────

def _build_property_analysis(wb, kpis, portfolio_name, eco_occ_target):
    ws = wb.create_sheet("Property Analysis")
    props = {k.property_name for k in kpis if not k.is_carveout}
    num_props = len(props)

    headers = [
        "Property", "Property Manager", "Period", "Total Units",
        "Actual Income", "Budget Income", "Income Variance", "Income Variance %",
        "Actual Expenses", "Budget Expenses", "Expense Variance", "Expense Variance %",
        "Actual NOI", "Budget NOI", "NOI Variance", "NOI Variance %",
        "Top NOI Driver 1", "Top NOI Driver 2",
        "Eco Occ %", "Budget Eco Occ %", "Eco Occ Variance",
        "GPR", "Vacancy", "Concessions", "Bad Debt",
        "Physical Occ %", "Leakage Gap",
        "Income/Unit", "Expense/Unit", "NOI/Unit",
        "Below Eco Occ Target?",
    ]
    ws.cell(1, 1, f"{portfolio_name} — Property Analysis ({num_props} Properties)").font = BOLD_FONT
    for col_idx, hdr in enumerate(headers, 1):
        cell = ws.cell(2, col_idx, hdr)
        add_kpi_comment(cell, hdr)
    style_header_row(ws, 2, len(headers))

    quarters = _get_sorted_quarters(kpis)
    row = 3
    for (yr, q) in quarters:
        period_lbl = _quarter_label(yr, q)
        q_kpis = _kpis_for_quarter(kpis, yr, q)

        prop_groups: dict[str, list] = {}
        for k in q_kpis:
            prop_groups.setdefault(k.property_name, []).append(k)

        for prop_name in sorted(prop_groups):
            group = prop_groups[prop_name]
            agg = _aggregate(group)
            pm = group[0].pm_name
            is_below = (agg.get("eco_occ_pct") or 0) < eco_occ_target and agg.get("eco_occ_pct") is not None
            row_data = [
                prop_name, pm, period_lbl, agg.get("total_units") or "Not Available",
                agg.get("actual_income"), agg.get("budget_income"),
                agg.get("income_variance"), agg.get("income_variance_pct"),
                agg.get("actual_expenses"), agg.get("budget_expenses"),
                agg.get("expense_variance"), agg.get("expense_variance_pct"),
                agg.get("actual_noi"), agg.get("budget_noi"),
                agg.get("noi_variance"), agg.get("noi_variance_pct"),
                group[0].top_noi_driver_1, group[0].top_noi_driver_2,
                agg.get("eco_occ_pct"), agg.get("budget_eco_occ_pct"), agg.get("eco_occ_variance"),
                agg.get("gpr"), agg.get("vacancy"), agg.get("concessions"), agg.get("bad_debt"),
                agg.get("physical_occ_pct"), agg.get("leakage_gap"),
                agg.get("income_per_unit"), agg.get("expense_per_unit"), agg.get("noi_per_unit"),
                "YES" if is_below else "no",
            ]
            for col_idx, val in enumerate(row_data, 1):
                ws.cell(row, col_idx, val)
            _apply_property_row_formats(ws, row, headers, agg)
            row += 1

    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"
    _autofit_columns(ws)


# ─── Property Monthly KPIs ────────────────────────────────────────────────────

def _build_monthly_kpis(wb, kpis):
    ws = wb.create_sheet("Property Monthly KPIs")
    headers = [
        "Property", "Property Manager", "Year", "Month", "Period",
        "Total Units",
        "Actual Income", "Budget Income", "Income Variance",
        "Actual Expenses", "Budget Expenses", "Expense Variance",
        "Actual NOI", "Budget NOI", "NOI Variance", "NOI Variance %",
        "GPR", "Vacancy", "Concessions", "Bad Debt",
        "Net Collectible", "Eco Occ %",
        "Physical Occ %", "Leakage Gap",
        "YoY Physical Occ Var", "YoY Eco Occ Var", "YoY Leakage Gap Change",
        "Income/Unit", "Expense/Unit", "NOI/Unit",
        "Source Key",
    ]
    for col_idx, hdr in enumerate(headers, 1):
        cell = ws.cell(1, col_idx, hdr)
        add_kpi_comment(cell, hdr)
    style_header_row(ws, 1, len(headers))

    row = 2
    for k in sorted(kpis, key=lambda x: (x.property_name, x.year, x.month)):
        ws.cell(row, 1,  k.property_name)
        ws.cell(row, 2,  k.pm_name)
        ws.cell(row, 3,  k.year)
        ws.cell(row, 4,  k.month)
        ws.cell(row, 5,  k.period)
        ws.cell(row, 6,  k.total_units if k.total_units is not None else "N/A")
        _c(ws, row, 7,  k.actual_income,   CURRENCY_FMT)
        _c(ws, row, 8,  k.budget_income,   CURRENCY_FMT)
        _c(ws, row, 9,  k.income_variance, CURRENCY_FMT)
        _c(ws, row, 10, k.actual_expenses,  CURRENCY_FMT)
        _c(ws, row, 11, k.budget_expenses,  CURRENCY_FMT)
        _c(ws, row, 12, k.expense_variance, CURRENCY_FMT)
        _c(ws, row, 13, k.actual_noi,       CURRENCY_FMT)
        _c(ws, row, 14, k.budget_noi,       CURRENCY_FMT)
        _c(ws, row, 15, k.noi_variance,     CURRENCY_FMT)
        _c(ws, row, 16, k.noi_variance_pct, VAR_PCT_FMT)
        _c(ws, row, 17, k.gpr,              CURRENCY_FMT)
        _c(ws, row, 18, k.vacancy,          CURRENCY_FMT)
        _c(ws, row, 19, k.concessions,      CURRENCY_FMT)
        _c(ws, row, 20, k.bad_debt,         CURRENCY_FMT)
        _c(ws, row, 21, k.net_collectible,  CURRENCY_FMT)
        _c(ws, row, 22, k.eco_occ_pct,      PCT_FMT)
        _c(ws, row, 23, k.physical_occ_pct, PCT_FMT)
        _c(ws, row, 24, k.leakage_gap,      PCT_FMT)
        _c(ws, row, 25, k.yoy_physical_occ_variance, PCT_FMT)
        _c(ws, row, 26, k.yoy_eco_occ_variance,      PCT_FMT)
        _c(ws, row, 27, k.yoy_leakage_gap_change,    PCT_FMT)
        _c(ws, row, 28, k.income_per_unit,  CURRENCY_FMT)
        _c(ws, row, 29, k.expense_per_unit, CURRENCY_FMT)
        _c(ws, row, 30, k.noi_per_unit,     CURRENCY_FMT)
        ws.cell(row, 31, k.source_key)
        row += 1

    ws.freeze_panes = "F2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    _autofit_columns(ws)


def _c(ws, row, col, value, fmt):
    """Write a cell value and apply a number format."""
    ws.cell(row, col, value).number_format = fmt


# ─── Aggregation ─────────────────────────────────────────────────────────────

def _aggregate(kpis: list[PropertyPeriodKPIs]) -> dict:
    """
    Aggregate numeric KPI fields across a list of PropertyPeriodKPIs records.

    Financial flows (income, expenses, GPR, etc.) are SUMMED — monthly amounts
    accumulate to a period total.

    Rate metrics (physical_occ_pct, eco_occ_pct) use a WEIGHTED approach:
      physical_occ = Σ(occupied_units) / Σ(total_units) across matched months.
    This avoids the >100% error that results from summing occupied_units and
    dividing by a single month's total_units.
    """
    def _sum(field):
        vals = [getattr(k, field) for k in kpis if getattr(k, field) is not None]
        return sum(vals) if vals else None

    actual_income   = _sum("actual_income")
    budget_income   = _sum("budget_income")
    actual_expenses = _sum("actual_expenses")
    budget_expenses = _sum("budget_expenses")

    actual_noi = (actual_income - actual_expenses) if (actual_income is not None and actual_expenses is not None) else None
    budget_noi = (budget_income - budget_expenses) if (budget_income is not None and budget_expenses is not None) else None

    gpr         = _sum("gpr")
    vacancy     = _sum("vacancy")
    concessions = _sum("concessions")
    bad_debt    = _sum("bad_debt")
    net_coll    = (gpr - (vacancy or 0) - (concessions or 0) - (bad_debt or 0)) if gpr is not None else None
    eco_occ     = (net_coll / gpr) if (net_coll is not None and gpr) else None

    # Budget eco occ: average the per-month budget rates across the period
    bud_eco_vals = [k.budget_eco_occ_pct for k in kpis if k.budget_eco_occ_pct is not None]
    bud_eco      = sum(bud_eco_vals) / len(bud_eco_vals) if bud_eco_vals else None
    eco_occ_var  = (eco_occ - bud_eco) if (eco_occ is not None and bud_eco is not None) else None

    # Physical occ: Σ(occupied) / Σ(total) — correct time-weighted average.
    # Using max(total_units) would produce values > 100% when aggregating multiple months.
    _paired = [
        (k.occupied_units, k.total_units)
        for k in kpis
        if k.occupied_units is not None and k.total_units is not None
    ]
    if _paired:
        _occ_sum   = sum(p[0] for p in _paired)
        _total_sum = sum(p[1] for p in _paired)
        phys_occ   = _occ_sum / _total_sum if _total_sum > 0 else None
    else:
        phys_occ = None
    total_units = next((k.total_units for k in kpis if k.total_units is not None), None)

    def _safe_div(a, b):
        return (a - b) if (a is not None and b is not None) else None

    def _safe_pct(a, b):
        return (a / abs(b)) if (a is not None and b is not None and b != 0) else None

    income_var  = _safe_div(actual_income, budget_income)
    expense_var = _safe_div(actual_expenses, budget_expenses)
    noi_var     = _safe_div(actual_noi, budget_noi)

    income_pu  = (actual_income / total_units)   if (actual_income   is not None and total_units) else None
    expense_pu = (actual_expenses / total_units) if (actual_expenses is not None and total_units) else None
    noi_pu     = (actual_noi / total_units)      if (actual_noi      is not None and total_units) else None

    return dict(
        actual_income=actual_income, budget_income=budget_income,
        income_variance=income_var, income_variance_pct=_safe_pct(income_var, budget_income),
        actual_expenses=actual_expenses, budget_expenses=budget_expenses,
        expense_variance=expense_var, expense_variance_pct=_safe_pct(expense_var, budget_expenses),
        actual_noi=actual_noi, budget_noi=budget_noi,
        noi_variance=noi_var, noi_variance_pct=_safe_pct(noi_var, budget_noi),
        gpr=gpr, vacancy=vacancy, concessions=concessions, bad_debt=bad_debt,
        net_collectible=net_coll, eco_occ_pct=eco_occ, budget_eco_occ_pct=bud_eco,
        eco_occ_variance=eco_occ_var,
        total_units=total_units, physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ) if (phys_occ is not None and eco_occ is not None) else None,
        income_per_unit=income_pu, expense_per_unit=expense_pu, noi_per_unit=noi_pu,
    )


# ─── Dashboard section helpers ────────────────────────────────────────────────

def _write_top_noi_table(ws, kpis, start_row, top_n, favorable):
    headers = ["Rank", "Property", "PM", "Prior Year NOI", "Current Year NOI",
               "NOI Variance", "NOI Variance %", "Driver 1", "Driver 2", "Commentary"]
    for col, h in enumerate(headers, 1):
        ws.cell(start_row, col, h)
    style_header_row(ws, start_row, len(headers), fill=SUBHDR_FILL, font=SUBHEADER_FONT)
    row = start_row + 1

    years = sorted({k.year for k in kpis})
    if len(years) < 2:
        ws.cell(row, 1, "Insufficient years for YoY comparison")
        return row + 1

    curr_yr, prev_yr = years[-1], years[-2]

    curr_by_prop: dict[str, float] = {}
    prev_by_prop: dict[str, float] = {}
    meta_by_prop: dict[str, PropertyPeriodKPIs] = {}
    for k in kpis:
        if k.actual_noi is None:
            continue
        if k.year == curr_yr:
            curr_by_prop[k.property_name] = curr_by_prop.get(k.property_name, 0) + k.actual_noi
            meta_by_prop[k.property_name] = k
        elif k.year == prev_yr:
            prev_by_prop[k.property_name] = prev_by_prop.get(k.property_name, 0) + k.actual_noi

    variances = []
    for prop in curr_by_prop:
        if prop in prev_by_prop:
            var = curr_by_prop[prop] - prev_by_prop[prop]
            variances.append((var, prop))

    variances.sort(key=lambda x: x[0], reverse=favorable)
    for rank, (var, prop) in enumerate(variances[:top_n], 1):
        prev_noi = prev_by_prop[prop]
        curr_noi = curr_by_prop[prop]
        var_pct = var / abs(prev_noi) if prev_noi else None
        k = meta_by_prop[prop]
        ws.cell(row, 1, rank)
        ws.cell(row, 2, prop)
        ws.cell(row, 3, k.pm_name)
        _c(ws, row, 4, prev_noi, CURRENCY_FMT)
        _c(ws, row, 5, curr_noi, CURRENCY_FMT)
        _c(ws, row, 6, var, CURRENCY_FMT)
        _c(ws, row, 7, var_pct, VAR_PCT_FMT)
        ws.cell(row, 8, k.top_noi_driver_1)
        ws.cell(row, 9, k.top_noi_driver_2)
        ws.cell(row, 10, k.commentary)
        row += 1
    return row


def _write_below_target_table(ws, kpis, start_row, eco_occ_target):
    headers = ["Property", "PM", "Eco Occ %", "Target", "Variance to Target",
               "Driver 1", "Driver 2", "Commentary"]
    for col, h in enumerate(headers, 1):
        ws.cell(start_row, col, h)
    style_header_row(ws, start_row, len(headers), fill=SUBHDR_FILL, font=SUBHEADER_FONT)
    row = start_row + 1

    below = [k for k in kpis if k.eco_occ_pct is not None and k.eco_occ_pct < eco_occ_target]
    below.sort(key=lambda k: k.eco_occ_pct or 1)
    for k in below:
        ws.cell(row, 1, k.property_name)
        ws.cell(row, 2, k.pm_name)
        _c(ws, row, 3, k.eco_occ_pct, PCT_FMT)
        _c(ws, row, 4, eco_occ_target, PCT_FMT)
        _c(ws, row, 5, (k.eco_occ_pct or 0) - eco_occ_target, PCT_FMT)
        ws.cell(row, 6, k.top_eco_occ_driver_1)
        ws.cell(row, 7, k.top_eco_occ_driver_2)
        ws.cell(row, 8, k.commentary)
        row += 1
    if row == start_row + 1:
        ws.cell(row, 1, f"All properties at or above {eco_occ_target:.0%} target")
    return row


def _generate_noi_commentary(kpis: list[PropertyPeriodKPIs]) -> str:
    """Rule-based template commentary for portfolio NOI trend."""
    years = sorted({k.year for k in kpis})
    if len(years) < 2:
        yr = years[0] if years else "N/A"
        total_noi = sum(k.actual_noi or 0 for k in kpis)
        return (
            f"Portfolio NOI for {yr} totaled ${total_noi:,.0f}. "
            "Prior year data not available for trend comparison."
        )

    curr_yr, prev_yr = years[-1], years[-2]
    curr_noi = sum(k.actual_noi or 0 for k in kpis if k.year == curr_yr)
    prev_noi = sum(k.actual_noi or 0 for k in kpis if k.year == prev_yr)
    curr_inc = sum(k.actual_income or 0 for k in kpis if k.year == curr_yr)
    prev_inc = sum(k.actual_income or 0 for k in kpis if k.year == prev_yr)
    curr_exp = sum(k.actual_expenses or 0 for k in kpis if k.year == curr_yr)
    prev_exp = sum(k.actual_expenses or 0 for k in kpis if k.year == prev_yr)

    noi_var = curr_noi - prev_noi
    noi_dir = "improved" if noi_var >= 0 else "declined"
    noi_pct = noi_var / abs(prev_noi) * 100 if prev_noi else 0
    inc_var = curr_inc - prev_inc
    exp_var = curr_exp - prev_exp
    primary_driver = "income" if abs(inc_var) >= abs(exp_var) else "expenses"

    curr_by_prop = {k.property_name: sum(x.actual_noi or 0 for x in kpis if x.year == curr_yr and x.property_name == k.property_name) for k in kpis if k.year == curr_yr}
    prev_by_prop = {k.property_name: sum(x.actual_noi or 0 for x in kpis if x.year == prev_yr and x.property_name == k.property_name) for k in kpis if k.year == prev_yr}
    prop_vars = {p: curr_by_prop.get(p, 0) - prev_by_prop.get(p, 0) for p in curr_by_prop}
    outlier = max(prop_vars, key=lambda p: abs(prop_vars[p])) if prop_vars else None
    outlier_share = abs(prop_vars[outlier]) / abs(noi_var) * 100 if (outlier and noi_var) else 0

    commentary = (
        f"Portfolio NOI {noi_dir} by ${abs(noi_var):,.0f} ({abs(noi_pct):.1f}%) "
        f"from {prev_yr} to {curr_yr}, primarily driven by "
        f"{'favorable' if inc_var >= 0 else 'unfavorable'} {primary_driver} trends. "
    )
    if outlier and outlier_share > 20:
        commentary += (
            f"{outlier} accounted for {outlier_share:.0f}% of the total NOI variance "
            f"(${prop_vars[outlier]:+,.0f}), suggesting this property-specific trend "
            "materially impacted portfolio results."
        )
    else:
        commentary += (
            "The variance appears broadly distributed across the portfolio "
            "with no single property dominating the trend."
        )
    return commentary


def _apply_property_row_formats(ws, row, headers, agg):
    fmt_map = {
        "Actual Income": CURRENCY_FMT, "Budget Income": CURRENCY_FMT,
        "Income Variance": CURRENCY_FMT, "Income Variance %": VAR_PCT_FMT,
        "Actual Expenses": CURRENCY_FMT, "Budget Expenses": CURRENCY_FMT,
        "Expense Variance": CURRENCY_FMT, "Expense Variance %": VAR_PCT_FMT,
        "Actual NOI": CURRENCY_FMT, "Budget NOI": CURRENCY_FMT,
        "NOI Variance": CURRENCY_FMT, "NOI Variance %": VAR_PCT_FMT,
        "Eco Occ %": PCT_FMT, "Budget Eco Occ %": PCT_FMT, "Eco Occ Variance": PCT_FMT,
        "Physical Occ %": PCT_FMT, "Leakage Gap": PCT_FMT,
        "GPR": CURRENCY_FMT, "Vacancy": CURRENCY_FMT,
        "Concessions": CURRENCY_FMT, "Bad Debt": CURRENCY_FMT,
        "Income/Unit": CURRENCY_FMT, "Expense/Unit": CURRENCY_FMT, "NOI/Unit": CURRENCY_FMT,
    }
    for col_idx, hdr in enumerate(headers, 1):
        fmt = fmt_map.get(hdr)
        if fmt:
            ws.cell(row, col_idx).number_format = fmt


def _autofit_columns(ws, max_width=60):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                length = len(str(cell.value or ""))
                if length > max_len:
                    max_len = length
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_len + 4, max_width)
