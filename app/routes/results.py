from flask import Blueprint, render_template, abort
from app.storage.runs import load_run
from config import ECO_OCC_TARGET

bp = Blueprint("results", __name__)

# Tooltip text for KPI labels in the web dashboard
_KPI_TOOLTIPS = {
    "Actual Income":      "Total Income = GPR + Other Income − Vacancy − Concessions − Bad Debt (Effective Gross Income)",
    "Budget Income":      "Budgeted Total Income for the period",
    "Income Variance":    "Actual Income − Budget Income. Positive = favorable (above budget)",
    "Income Variance %":  "Income Variance / Budget Income",
    "Actual Expenses":    "Sum of all Operating Expense accounts. Excludes depreciation, debt service, reserves",
    "Budget Expenses":    "Budgeted Operating Expenses for the period",
    "Expense Variance":   "Actual Expenses − Budget Expenses. Negative = favorable (under budget)",
    "Expense Variance %": "Expense Variance / Budget Expenses",
    "Actual NOI":         "NOI = Total Income − Total Operating Expenses",
    "Budget NOI":         "Budget NOI = Budget Income − Budget Expenses",
    "NOI Variance":       "NOI Variance = Actual NOI − Budget NOI. Positive = favorable",
    "NOI Variance %":     "NOI Variance / |Budget NOI|. Absolute denominator handles sign flips when budget NOI is negative",
    "GPR":                "Gross Potential Rent — total scheduled rent before any deductions",
    "Vacancy":            "Vacancy loss — rent foregone from unoccupied units",
    "Concessions":        "Move-in specials and rent concessions",
    "Bad Debt":           "Collection losses and write-offs",
    "Net Collectible":    "GPR − Vacancy − Concessions − Bad Debt",
    "Eco Occ %":          "Economic Occupancy % = Net Collectible / GPR",
    "Budget Eco Occ %":   "Budget Economic Occupancy % = Budget Net Collectible / Budget GPR",
    "Eco Occ Variance":   "Actual Eco Occ % − Budget Eco Occ %",
    "Physical Occ %":     "Physical Occ % = Occupied Units / Total Units. Sourced from Physical Occupancy Report",
    "Leakage Gap":        "Physical Occ % − Economic Occ %. Positive = units occupied but rent not being fully collected",
    "Income/Unit":        "Actual Income / Total Units (from Physical Occupancy Report)",
    "Expense/Unit":       "Actual Expenses / Total Units (from Physical Occupancy Report)",
    "NOI/Unit":           "Actual NOI / Total Units (from Physical Occupancy Report)",
}

# KPI row definitions for the transposed Portfolio Summary table.
# Each entry is either:
#   None  → separator row
#   (label, agg_dict_key, fmt, favorable_positive)
#     favorable_positive:
#       True  = green when positive (income variances, NOI variances)
#       False = green when negative (expense variances, leakage gap)
#       None  = no color applied
_SUMMARY_KPI_DEFINITIONS = [
    ("Actual Income",      "actual_income",        "currency", None),
    ("Budget Income",      "budget_income",        "currency", None),
    ("Income Variance",    "income_variance",      "currency", True),
    ("Income Variance %",  "income_variance_pct",  "pct",      True),
    None,
    ("Actual Expenses",    "actual_expenses",      "currency", None),
    ("Budget Expenses",    "budget_expenses",      "currency", None),
    ("Expense Variance",   "expense_variance",     "currency", False),
    ("Expense Variance %", "expense_variance_pct", "pct",      False),
    None,
    ("Actual NOI",         "actual_noi",           "currency", None),
    ("Budget NOI",         "budget_noi",           "currency", None),
    ("NOI Variance",       "noi_variance",         "currency", True),
    ("NOI Variance %",     "noi_variance_pct",     "pct",      True),
    None,
    ("GPR",                "gpr",                  "currency", None),
    ("Vacancy",            "vacancy",              "currency", None),
    ("Concessions",        "concessions",          "currency", None),
    ("Bad Debt",           "bad_debt",             "currency", None),
    ("Net Collectible",    "net_collectible",      "currency", None),
    ("Eco Occ %",          "eco_occ_pct",          "pct",      None),
    ("Budget Eco Occ %",   "budget_eco_occ_pct",   "pct",      None),
    ("Eco Occ Variance",   "eco_occ_variance",     "pct",      True),
    None,
    ("Physical Occ %",     "physical_occ_pct",     "pct",      None),
    ("Leakage Gap",        "leakage_gap",          "pct",      False),
    None,
    ("Income/Unit",        "income_per_unit",      "currency", None),
    ("Expense/Unit",       "expense_per_unit",     "currency", None),
    ("NOI/Unit",           "noi_per_unit",         "currency", None),
]


def _month_to_quarter(month: int) -> int:
    return (month - 1) // 3 + 1


def _quarter_label(year: int, quarter: int) -> str:
    return f"Q{quarter} - {year}"


def _agg_kpis(kpi_dicts: list[dict]) -> dict:
    """Aggregate a list of KPI dicts (loaded from JSON) into a single summary dict."""
    def _sum(field):
        vals = [k[field] for k in kpi_dicts if k.get(field) is not None]
        return sum(vals) if vals else None

    actual_income   = _sum("actual_income")
    budget_income   = _sum("budget_income")
    actual_expenses = _sum("actual_expenses")
    budget_expenses = _sum("budget_expenses")
    gpr             = _sum("gpr")
    vacancy         = _sum("vacancy")
    concessions     = _sum("concessions")
    bad_debt        = _sum("bad_debt")

    actual_noi = (actual_income - actual_expenses) if (actual_income is not None and actual_expenses is not None) else None
    budget_noi = (budget_income - budget_expenses) if (budget_income is not None and budget_expenses is not None) else None

    net_coll = (gpr - (vacancy or 0) - (concessions or 0) - (bad_debt or 0)) if gpr is not None else None
    eco_occ  = (net_coll / gpr) if (net_coll is not None and gpr) else None

    # Budget eco occ: average the per-month budget rates across the period
    bud_eco_vals = [k["budget_eco_occ_pct"] for k in kpi_dicts if k.get("budget_eco_occ_pct") is not None]
    bud_eco      = sum(bud_eco_vals) / len(bud_eco_vals) if bud_eco_vals else None
    eco_occ_var  = (eco_occ - bud_eco) if (eco_occ is not None and bud_eco is not None) else None

    noi_var     = (actual_noi - budget_noi) if (actual_noi is not None and budget_noi is not None) else None
    noi_var_pct = (noi_var / abs(budget_noi)) if (noi_var is not None and budget_noi) else None

    # Physical occ: Σ(occupied_units) / Σ(total_units) — correct time-weighted average.
    # IMPORTANT: do NOT use max(total_units) with sum(occupied_units) — that produces
    # values > 100% when aggregating multiple monthly records (e.g. 3 months × 100 units
    # at 95% → sum(occ)=285, max(total)=100, ratio=285% — wrong).
    _paired = [
        (k["occupied_units"], k["total_units"])
        for k in kpi_dicts
        if k.get("occupied_units") is not None and k.get("total_units") is not None
    ]
    if _paired:
        _occ_sum   = sum(p[0] for p in _paired)
        _total_sum = sum(p[1] for p in _paired)
        phys_occ   = _occ_sum / _total_sum if _total_sum > 0 else None
    else:
        phys_occ = None

    # Total units: use the most common non-None value (first encountered is fine for display)
    total_units = next((k["total_units"] for k in kpi_dicts if k.get("total_units") is not None), None)

    income_pu  = (actual_income   / total_units) if (actual_income   is not None and total_units) else None
    expense_pu = (actual_expenses / total_units) if (actual_expenses is not None and total_units) else None
    noi_pu     = (actual_noi      / total_units) if (actual_noi      is not None and total_units) else None

    def _safe_pct(num, denom):
        if num is None or denom is None or denom == 0:
            return None
        return num / abs(denom)

    return dict(
        actual_income=actual_income,
        budget_income=budget_income,
        income_variance=(actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None,
        income_variance_pct=_safe_pct(
            (actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None,
            budget_income,
        ),
        actual_expenses=actual_expenses,
        budget_expenses=budget_expenses,
        expense_variance=(actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None,
        expense_variance_pct=_safe_pct(
            (actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None,
            budget_expenses,
        ),
        actual_noi=actual_noi,
        budget_noi=budget_noi,
        noi_variance=noi_var,
        noi_variance_pct=noi_var_pct,
        eco_occ_pct=eco_occ,
        budget_eco_occ_pct=bud_eco,
        eco_occ_variance=eco_occ_var,
        physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ) if (phys_occ is not None and eco_occ is not None) else None,
        income_per_unit=income_pu,
        expense_per_unit=expense_pu,
        noi_per_unit=noi_pu,
        gpr=gpr,
        vacancy=vacancy,
        concessions=concessions,
        bad_debt=bad_debt,
        net_collectible=net_coll,
        total_units=total_units,
    )


@bp.route("/results/<run_id>")
def show(run_id):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    meta   = data["metadata"]
    kpis   = data["kpis"]
    checks = data["quality_checks"]

    portfolio_name = meta.get("portfolio_name", "Portfolio")
    eco_occ_target = meta.get("eco_occ_target", ECO_OCC_TARGET)
    years  = meta.get("years", [])
    props  = meta.get("properties", [])

    # ── Quarter-period aggregation for the transposed Portfolio Summary table ──
    all_quarters: set[tuple] = set()
    for k in kpis:
        if not k.get("is_carveout") and k.get("year") and k.get("month"):
            all_quarters.add((k["year"], _month_to_quarter(k["month"])))
    sorted_quarters = sorted(all_quarters)
    period_labels = [_quarter_label(yr, q) for (yr, q) in sorted_quarters]

    # Aggregate all non-carveout KPIs for each quarter-period
    period_aggs: dict[str, dict] = {}
    for (yr, q) in sorted_quarters:
        months = {(q - 1) * 3 + 1, (q - 1) * 3 + 2, (q - 1) * 3 + 3}
        q_kpis = [
            k for k in kpis
            if k.get("year") == yr and k.get("month") in months and not k.get("is_carveout")
        ]
        period_aggs[_quarter_label(yr, q)] = _agg_kpis(q_kpis)

    # Build summary_kpi_rows: one entry per KPI definition row
    summary_kpi_rows = []
    for defn in _SUMMARY_KPI_DEFINITIONS:
        if defn is None:
            summary_kpi_rows.append({"sep": True})
            continue
        label, key, fmt, fav = defn
        values = [period_aggs.get(lbl, {}).get(key) for lbl in period_labels]
        summary_kpi_rows.append({
            "sep": False,
            "label": label,
            "key": key,
            "fmt": fmt,
            "favorable_positive": fav,
            "tooltip": _KPI_TOOLTIPS.get(label, ""),
            "period_values": values,
        })

    # ── Property table: aggregate each property across its latest year ──
    latest_yr = max(years) if years else None
    prop_rows = []
    for prop in sorted(props):
        prop_kpis = [
            k for k in kpis
            if k.get("property_name") == prop and k.get("year") == latest_yr
        ]
        if prop_kpis:
            agg = _agg_kpis(prop_kpis)
            agg["property_name"] = prop
            agg["pm_name"] = prop_kpis[0].get("pm_name", "")
            agg["top_noi_driver_1"] = prop_kpis[0].get("top_noi_driver_1", "")
            agg["top_noi_driver_2"] = prop_kpis[0].get("top_noi_driver_2", "")
            prop_rows.append(agg)

    return render_template(
        "dashboard.html",
        run_id=run_id,
        meta=meta,
        kpis=kpis,
        period_labels=period_labels,
        summary_kpi_rows=summary_kpi_rows,
        prop_rows=prop_rows,
        quality_checks=checks,
        portfolio_name=portfolio_name,
        eco_occ_target=eco_occ_target,
        years=years,
        properties=props,
        num_properties=len(props),
    )
