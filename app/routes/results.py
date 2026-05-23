from flask import Blueprint, render_template, abort
from app.storage.runs import load_run
from config import ECO_OCC_TARGET

bp = Blueprint("results", __name__)


def _agg_kpis(kpi_dicts: list[dict]) -> dict:
    """Aggregate a list of KPI dicts (loaded from JSON) into a summary dict."""
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
    net_coll   = (gpr - (vacancy or 0) - (concessions or 0) - (bad_debt or 0)) if gpr is not None else None
    eco_occ    = (net_coll / gpr) if (net_coll is not None and gpr) else None

    noi_var     = (actual_noi - budget_noi) if (actual_noi is not None and budget_noi is not None) else None
    noi_var_pct = (noi_var / abs(budget_noi)) if (noi_var is not None and budget_noi) else None

    total_units = max((k.get("total_units") or 0 for k in kpi_dicts), default=0) or None
    occ_units   = _sum("occupied_units")
    phys_occ    = (occ_units / total_units) if (occ_units is not None and total_units) else None

    income_pu  = (actual_income / total_units)   if (actual_income   is not None and total_units) else None
    expense_pu = (actual_expenses / total_units) if (actual_expenses is not None and total_units) else None
    noi_pu     = (actual_noi / total_units)      if (actual_noi      is not None and total_units) else None

    return dict(
        actual_income=actual_income, budget_income=budget_income,
        income_variance=(actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None,
        actual_expenses=actual_expenses, budget_expenses=budget_expenses,
        expense_variance=(actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None,
        actual_noi=actual_noi, budget_noi=budget_noi,
        noi_variance=noi_var, noi_variance_pct=noi_var_pct,
        eco_occ_pct=eco_occ, physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ) if (phys_occ is not None and eco_occ is not None) else None,
        income_per_unit=income_pu, expense_per_unit=expense_pu, noi_per_unit=noi_pu,
        gpr=gpr, vacancy=vacancy, concessions=concessions, bad_debt=bad_debt,
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

    summary_rows = []
    for yr in sorted(years):
        yr_kpis = [k for k in kpis if k.get("year") == yr and not k.get("is_carveout")]
        agg = _agg_kpis(yr_kpis)
        agg["year"] = yr
        summary_rows.append(agg)

    latest_yr = max(years) if years else None
    prop_rows = []
    for prop in sorted(props):
        prop_kpis = [k for k in kpis if k.get("property_name") == prop and k.get("year") == latest_yr]
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
        summary_rows=summary_rows,
        prop_rows=prop_rows,
        quality_checks=checks,
        portfolio_name=portfolio_name,
        eco_occ_target=eco_occ_target,
        years=years,
        properties=props,
        num_properties=len(props),
    )
