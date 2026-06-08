from flask import Blueprint, render_template, abort
from app.storage.runs import load_run

bp = Blueprint("property_detail", __name__)

_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def _ar_period_label(year: int, month: int) -> str:
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


def _compute_prop_projection(prop_kpis: list[dict]) -> tuple[str, dict]:
    """
    Compute full-year projection for a single property from its KPI dicts.
    Returns (proj_yr_label, projection_dict).

    projection_dict keys: "actual_income", "actual_expenses", "actual_noi"
    Each maps to: {q1_actual, proj_fy, fy_budget, var_to_plan, var_to_plan_pct}
    """
    if not prop_kpis:
        return "", {}

    years_present = {k["year"] for k in prop_kpis if k.get("year")}
    if not years_present:
        return "", {}
    proj_yr = max(years_present)
    proj_yr_label = str(proj_yr)

    def _psum(kpi_list, field):
        vals = [k.get(field) for k in kpi_list if k.get(field) is not None]
        return sum(vals) if vals else None

    q1k   = [k for k in prop_kpis if k.get("year") == proj_yr and k.get("month") in (1, 2, 3)]
    q2q4k = [k for k in prop_kpis if k.get("year") == proj_yr and k.get("month") in range(4, 13)]
    ayk   = [k for k in prop_kpis if k.get("year") == proj_yr]

    projection = {}
    for pk, bk in [("actual_income",   "budget_income"),
                   ("actual_expenses",  "budget_expenses"),
                   ("actual_noi",       "budget_noi")]:
        q1_act   = _psum(q1k,   pk)
        q2q4_bud = _psum(q2q4k, bk)
        ay_bud   = _psum(ayk,   bk)

        if not q2q4_bud:
            q1_bud   = _psum(q1k, bk)
            q2q4_bud = (q1_bud * 3) if q1_bud is not None else None
            fy_bud   = (q1_bud * 4) if q1_bud is not None else None
        else:
            fy_bud = ay_bud

        proj_fy = (q1_act + q2q4_bud) if (q1_act is not None and q2q4_bud is not None) else None
        var     = (proj_fy - fy_bud)   if (proj_fy is not None and fy_bud is not None) else None
        var_pct = (var / abs(fy_bud))  if (var is not None and fy_bud) else None

        projection[pk] = {
            "q1_actual":     q1_act,
            "proj_fy":       proj_fy,
            "fy_budget":     fy_bud,
            "var_to_plan":   var,
            "var_to_plan_pct": var_pct,
        }

    return proj_yr_label, projection


@bp.route("/results/<run_id>/property/<property_name>")
def show(run_id, property_name):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    prop_kpis = [k for k in data["kpis"] if k["property_name"] == property_name]
    if not prop_kpis:
        abort(404)
    # Newest period first
    prop_kpis.sort(key=lambda k: (-k["year"], -k["month"]))

    # Build AR aging detail for this property
    # AR rows are stored as dicts; @property fields (total_overdue, pct_overdue)
    # are not serialized by asdict() so recompute them here.
    raw_ar = data.get("ar_aging", [])
    prop_ar = [r for r in raw_ar if r["property_name"] == property_name]

    # Augment each row with computed fields and sort
    for r in prop_ar:
        r["total_overdue"] = r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"]
        charge = r.get("charge_amount", 0)
        r["pct_overdue"] = (r["total_overdue"] / charge) if charge and charge > 0 else None
        r["period_label"] = _ar_period_label(r["year"], r["month"])

    # Newest period first within each receivable type
    prop_ar.sort(key=lambda r: (r["receivable_type"], -r["year"], -r["month"]))

    # Split into sub-dicts by receivable type for template simplicity
    ar_by_type = {}
    for r in prop_ar:
        ar_by_type.setdefault(r["receivable_type"], []).append(r)

    # Compute per-property full-year projection for the latest year in this property's data
    proj_yr_label, prop_projection = _compute_prop_projection(prop_kpis)

    return render_template(
        "property_detail.html",
        run_id=run_id,
        property_name=property_name,
        kpis=prop_kpis,
        meta=data["metadata"],
        ar_by_type=ar_by_type,
        prop_projection=prop_projection,
        proj_yr_label=proj_yr_label,
    )
