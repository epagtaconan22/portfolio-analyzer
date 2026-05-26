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


# ── Quarter helpers ────────────────────────────────────────────────────────────

def _month_to_quarter(month: int) -> int:
    return (month - 1) // 3 + 1


def _quarter_label(year: int, quarter: int) -> str:
    return f"Q{quarter} - {year}"


# ── AR Aging helpers ───────────────────────────────────────────────────────────

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


def _ar_period_label(year: int, month: int) -> str:
    """Return display label e.g. 'Mar-2024'."""
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


def _agg_ar(ar_rows: list[dict], receivable_type: str, year: int, month: int) -> dict | None:
    """Aggregate AR rows for a specific type/period. Returns None if no matching rows."""
    rows = [r for r in ar_rows
            if r["receivable_type"] == receivable_type
            and r["year"] == year
            and r["month"] == month]
    if not rows:
        return None
    charge_amount = sum(r["charge_amount"] for r in rows)
    total_overdue = sum(r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed":   sum(r["current_owed"] for r in rows),
        "prepayments":    sum(r["prepayments"] for r in rows),
        "pct_overdue":    total_overdue / charge_amount if charge_amount > 0 else None,
        "property_count": len({r["property_name"] for r in rows}),
    }


def _agg_ar_for_prop(ar_rows: list[dict], property_name: str,
                     receivable_type: str, year: int, month: int) -> dict | None:
    """Aggregate AR rows for a specific property/type/period."""
    rows = [r for r in ar_rows
            if r["property_name"] == property_name
            and r["receivable_type"] == receivable_type
            and r["year"] == year
            and r["month"] == month]
    if not rows:
        return None
    charge = sum(r["charge_amount"] for r in rows)
    overdue = sum(r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed": sum(r["current_owed"] for r in rows),
        "prepayments":  sum(r["prepayments"] for r in rows),
        "pct_overdue":  overdue / charge if charge > 0 else None,
    }


def _ar_yoy_delta(curr: dict, prev: dict) -> dict:
    """Compute YoY delta between two period aggregates."""
    pct_delta = None
    if curr.get("pct_overdue") is not None and prev.get("pct_overdue") is not None:
        pct_delta = curr["pct_overdue"] - prev["pct_overdue"]
    return {
        "current_owed_delta": curr["current_owed"] - prev["current_owed"],
        "prepayments_delta":  curr["prepayments"] - prev["prepayments"],
        "pct_overdue_delta":  pct_delta,
    }


def _pct_delta(curr: dict | None, prev: dict | None) -> float | None:
    """Return pct_overdue delta between two property aggregates, or None."""
    if curr and prev:
        c = curr.get("pct_overdue")
        p = prev.get("pct_overdue")
        if c is not None and p is not None:
            return c - p
    return None


# ── KPI aggregation ────────────────────────────────────────────────────────────

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

    bud_eco_vals = [k["budget_eco_occ_pct"] for k in kpi_dicts if k.get("budget_eco_occ_pct") is not None]
    bud_eco      = sum(bud_eco_vals) / len(bud_eco_vals) if bud_eco_vals else None
    eco_occ_var  = (eco_occ - bud_eco) if (eco_occ is not None and bud_eco is not None) else None

    noi_var     = (actual_noi - budget_noi) if (actual_noi is not None and budget_noi is not None) else None
    noi_var_pct = (noi_var / abs(budget_noi)) if (noi_var is not None and budget_noi) else None

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


# ── Route ──────────────────────────────────────────────────────────────────────

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

    # ── Quarter-period aggregation ─────────────────────────────────────────────
    all_quarters: set[tuple] = set()
    for k in kpis:
        if not k.get("is_carveout") and k.get("year") and k.get("month"):
            all_quarters.add((k["year"], _month_to_quarter(k["month"])))
    sorted_quarters = sorted(all_quarters)
    period_labels = [_quarter_label(yr, q) for (yr, q) in sorted_quarters]

    period_aggs: dict[str, dict] = {}
    period_property_counts: dict[str, int] = {}
    for (yr, q) in sorted_quarters:
        months = {(q - 1) * 3 + 1, (q - 1) * 3 + 2, (q - 1) * 3 + 3}
        q_kpis = [
            k for k in kpis
            if k.get("year") == yr and k.get("month") in months and not k.get("is_carveout")
        ]
        lbl = _quarter_label(yr, q)
        period_aggs[lbl] = _agg_kpis(q_kpis)
        period_property_counts[lbl] = len({k["property_name"] for k in q_kpis})

    latest_period_label = period_labels[-1] if period_labels else ""

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

    # ── Property table ─────────────────────────────────────────────────────────
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

    # ── AR Aging section ───────────────────────────────────────────────────────
    ar_rows = data.get("ar_aging", [])
    ar_summary: dict = {}
    ar_prop_rows: list = []
    ar_latest_period_label: str = ""

    if ar_rows:
        for rtype in ["Tenant Rent", "Subsidy"]:
            periods = sorted({(r["year"], r["month"]) for r in ar_rows
                              if r["receivable_type"] == rtype})
            if not periods:
                continue
            period_set = set(periods)
            cols = []
            for (yr, mo) in periods:
                agg = _agg_ar(ar_rows, rtype, yr, mo)
                cols.append({
                    "type":         "period",
                    "label":        _ar_period_label(yr, mo),
                    "year": yr, "month": mo,
                    "current_owed": agg["current_owed"] if agg else None,
                    "prepayments":  agg["prepayments"]  if agg else None,
                    "pct_overdue":  agg["pct_overdue"]  if agg else None,
                })
                if (yr - 1, mo) in period_set:
                    curr = agg
                    prev = _agg_ar(ar_rows, rtype, yr - 1, mo)
                    if curr and prev:
                        delta = _ar_yoy_delta(curr, prev)
                    else:
                        delta = {"current_owed_delta": None,
                                 "prepayments_delta": None,
                                 "pct_overdue_delta": None}
                    cols.append({
                        "type":         "yoy",
                        "label":        "YoY Δ",
                        "year": yr, "month": mo,
                        "current_owed": delta["current_owed_delta"],
                        "prepayments":  delta["prepayments_delta"],
                        "pct_overdue":  delta["pct_overdue_delta"],
                    })

            prop_count = len({r["property_name"] for r in ar_rows
                              if r["receivable_type"] == rtype})
            ar_summary[rtype] = {"property_count": prop_count, "cols": cols}

        # Property-level AR table — latest period across all types
        latest_yr_ar, latest_mo_ar = max((r["year"], r["month"]) for r in ar_rows)
        ar_latest_period_label = _ar_period_label(latest_yr_ar, latest_mo_ar)

        all_latest_props = sorted({r["property_name"] for r in ar_rows
                                   if r["year"] == latest_yr_ar and r["month"] == latest_mo_ar})

        for prop in all_latest_props:
            pm = next((r["pm_name"] for r in ar_rows
                       if r["property_name"] == prop
                       and r["year"] == latest_yr_ar
                       and r["month"] == latest_mo_ar), "")

            tr_curr  = _agg_ar_for_prop(ar_rows, prop, "Tenant Rent", latest_yr_ar, latest_mo_ar)
            tr_prev  = _agg_ar_for_prop(ar_rows, prop, "Tenant Rent", latest_yr_ar - 1, latest_mo_ar)
            sub_curr = _agg_ar_for_prop(ar_rows, prop, "Subsidy",     latest_yr_ar, latest_mo_ar)
            sub_prev = _agg_ar_for_prop(ar_rows, prop, "Subsidy",     latest_yr_ar - 1, latest_mo_ar)

            ar_prop_rows.append({
                "property_name":   prop,
                "pm_name":         pm,
                "tr_current_owed":  tr_curr["current_owed"]  if tr_curr else None,
                "tr_prepayments":   tr_curr["prepayments"]   if tr_curr else None,
                "tr_pct_overdue":   tr_curr["pct_overdue"]   if tr_curr else None,
                "tr_yoy_pct_delta": _pct_delta(tr_curr, tr_prev),
                "sub_current_owed":  sub_curr["current_owed"]  if sub_curr else None,
                "sub_prepayments":   sub_curr["prepayments"]   if sub_curr else None,
                "sub_pct_overdue":   sub_curr["pct_overdue"]   if sub_curr else None,
                "sub_yoy_pct_delta": _pct_delta(sub_curr, sub_prev),
            })

        # Sort by Tenant Rent current_owed descending
        ar_prop_rows.sort(key=lambda r: (r["tr_current_owed"] or 0), reverse=True)

    return render_template(
        "dashboard.html",
        run_id=run_id,
        meta=meta,
        kpis=kpis,
        period_labels=period_labels,
        period_property_counts=period_property_counts,
        summary_kpi_rows=summary_kpi_rows,
        prop_rows=prop_rows,
        quality_checks=checks,
        portfolio_name=portfolio_name,
        eco_occ_target=eco_occ_target,
        latest_period_label=latest_period_label,
        years=years,
        properties=props,
        num_properties=len(props),
        # AR Aging context
        ar_summary=ar_summary,
        ar_prop_rows=ar_prop_rows,
        ar_latest_period_label=ar_latest_period_label,
    )
