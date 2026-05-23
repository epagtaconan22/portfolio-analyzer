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


def _build_dashboard(wb, kpis, portfolio_name, eco_occ_target):
    ws = wb.create_sheet("Dashboard")

    props = {k.property_name for k in kpis if not k.is_carveout}
    num_props = len(props)

    row = 1
    ws.cell(row, 1, f"{portfolio_name} — Portfolio Summary ({num_props} Properties)")
    ws.cell(row, 1).font = BOLD_FONT
    row += 2

    summary_headers = [
        "Period", "Actual Income", "Budget Income", "Income Variance", "Income Variance %",
        "Actual Expenses", "Budget Expenses", "Expense Variance", "Expense Variance %",
        "Actual NOI", "Budget NOI", "NOI Variance", "NOI Variance %",
        "Eco Occ %", "Budget Eco Occ %", "Eco Occ Variance",
        "Physical Occ %", "Leakage Gap",
        "Income/Unit", "Expense/Unit", "NOI/Unit",
    ]
    for col_idx, hdr in enumerate(summary_headers, 1):
        cell = ws.cell(row, col_idx, hdr)
        add_kpi_comment(cell, hdr)
    style_header_row(ws, row, len(summary_headers))
    row += 1

    years = sorted({k.year for k in kpis})
    for yr in years:
        yr_kpis = [k for k in kpis if k.year == yr and not k.is_carveout]
        agg = _aggregate(yr_kpis)
        _write_summary_row(ws, row, str(yr), agg, eco_occ_target)
        row += 1

    row += 2

    ws.cell(row, 1, "NOI Trend Commentary").font = BOLD_FONT
    row += 1
    commentary = _generate_noi_commentary(kpis)
    ws.cell(row, 1, commentary)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 120
    row += 3

    ws.cell(row, 1, f"Top Positive NOI Variances ({num_props} Properties Analyzed)").font = BOLD_FONT
    row += 1
    row = _write_top_noi_table(ws, kpis, row, top_n=5, favorable=True)
    row += 2

    ws.cell(row, 1, f"Top Negative NOI Variances ({num_props} Properties Analyzed)").font = BOLD_FONT
    row += 1
    row = _write_top_noi_table(ws, kpis, row, top_n=5, favorable=False)
    row += 2

    ws.cell(row, 1, f"Properties Below Economic Occupancy Target ({eco_occ_target:.0%})").font = BOLD_FONT
    row += 1
    _write_below_target_table(ws, kpis, row, eco_occ_target)

    ws.freeze_panes = "B4"
    _autofit_columns(ws)


def _build_property_analysis(wb, kpis, portfolio_name, eco_occ_target):
    ws = wb.create_sheet("Property Analysis")
    props = {k.property_name for k in kpis if not k.is_carveout}
    num_props = len(props)

    headers = [
        "Property", "Property Manager", "Total Units",
        "Actual Income", "Budget Income", "Income Variance", "Income Variance %",
        "Actual Expenses", "Budget Expenses", "Expense Variance", "Expense Variance %",
        "Actual NOI", "Budget NOI", "NOI Variance", "NOI Variance %",
        "Top NOI Driver 1", "Top NOI Driver 2",
        "Eco Occ %", "Budget Eco Occ %", "Eco Occ Variance",
        "GPR", "Vacancy", "Concessions", "Bad Debt",
        "Physical Occ %", "Leakage Gap",
        "Income/Unit", "Expense/Unit", "NOI/Unit",
        "Below Eco Occ Target?",
        "Commentary",
    ]
    ws.cell(1, 1, f"{portfolio_name} — Property Analysis ({num_props} Properties)").font = BOLD_FONT
    for col_idx, hdr in enumerate(headers, 1):
        cell = ws.cell(2, col_idx, hdr)
        add_kpi_comment(cell, hdr)
    style_header_row(ws, 2, len(headers))

    years = sorted({k.year for k in kpis})
    row = 3
    for yr in years:
        yr_kpis = [k for k in kpis if k.year == yr]
        prop_groups: dict[str, list] = {}
        for k in yr_kpis:
            prop_groups.setdefault(k.property_name, []).append(k)

        for prop_name in sorted(prop_groups):
            group = prop_groups[prop_name]
            agg = _aggregate(group)
            pm = group[0].pm_name
            is_below = (agg.get("eco_occ_pct") or 0) < eco_occ_target and agg.get("eco_occ_pct") is not None
            row_data = [
                prop_name, pm, agg.get("total_units") or "Not Available",
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
                group[0].commentary,
            ]
            for col_idx, val in enumerate(row_data, 1):
                ws.cell(row, col_idx, val)
            _apply_property_row_formats(ws, row, headers, agg)
            row += 1

    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}2"
    _autofit_columns(ws)


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
    """Sum/weight numeric KPI fields across a list of KPI records."""
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

    # Budget eco occ: weighted sum of budget components
    bud_gpr     = _sum("gpr")  # approximation; full budget path uses budget-specific rows
    bud_eco     = None

    total_units = max((k.total_units or 0 for k in kpis), default=0) or None
    occ_units   = _sum("occupied_units")
    phys_occ    = (occ_units / total_units) if (occ_units is not None and total_units) else None

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
        eco_occ_variance=None,
        total_units=total_units, occupied_units=occ_units, physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ) if (phys_occ is not None and eco_occ is not None) else None,
        income_per_unit=income_pu, expense_per_unit=expense_pu, noi_per_unit=noi_pu,
    )


# ─── Dashboard helpers ───────────────────────────────────────────────────────

def _write_summary_row(ws, row, label, agg, eco_occ_target):
    ws.cell(row, 1, label)
    vals = [
        agg.get("actual_income"), agg.get("budget_income"),
        agg.get("income_variance"), agg.get("income_variance_pct"),
        agg.get("actual_expenses"), agg.get("budget_expenses"),
        agg.get("expense_variance"), agg.get("expense_variance_pct"),
        agg.get("actual_noi"), agg.get("budget_noi"),
        agg.get("noi_variance"), agg.get("noi_variance_pct"),
        agg.get("eco_occ_pct"), agg.get("budget_eco_occ_pct"), agg.get("eco_occ_variance"),
        agg.get("physical_occ_pct"), agg.get("leakage_gap"),
        agg.get("income_per_unit"), agg.get("expense_per_unit"), agg.get("noi_per_unit"),
    ]
    currency_cols = {2, 3, 4, 6, 7, 8, 10, 11, 12, 19, 20, 21}
    pct_cols      = {5, 9, 13, 14, 15, 16, 17}
    for i, val in enumerate(vals, 2):
        cell = ws.cell(row, i, val)
        if i in currency_cols:
            cell.number_format = CURRENCY_FMT
        elif i in pct_cols:
            cell.number_format = PCT_FMT


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

    # Aggregate each property's total NOI per year
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
