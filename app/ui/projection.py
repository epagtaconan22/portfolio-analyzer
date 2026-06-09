"""Full-year projection helper — extracted from routes/property_detail.py."""


def compute_prop_projection(prop_kpis: list[dict]) -> tuple[str, dict]:
    """
    Compute full-year projection for a single property's KPI list.
    Returns (proj_yr_label, projection_dict).

    projection_dict keys: "actual_income", "actual_expenses", "actual_noi"
    Each value is a dict: {q1_actual, proj_fy, fy_budget, var_to_plan, var_to_plan_pct}
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
            q1_bud = _psum(q1k, bk)
            q1_bud_per_month = (q1_bud / 3) if q1_bud is not None else None
            q2q4_bud = (q1_bud_per_month * 3) if q1_bud_per_month is not None else None
            fy_bud = (q1_bud_per_month * 4) if q1_bud_per_month is not None else None
        else:
            fy_bud = ay_bud

        proj_fy = (q1_act + q2q4_bud) if (q1_act is not None and q2q4_bud is not None) else None
        var     = (proj_fy - fy_bud)   if (proj_fy is not None and fy_bud is not None) else None
        var_pct = (var / abs(fy_bud))  if (var is not None and fy_bud) else None

        projection[pk] = {
            "q1_actual":       q1_act,
            "proj_fy":         proj_fy,
            "fy_budget":       fy_bud,
            "var_to_plan":     var,
            "var_to_plan_pct": var_pct,
        }

    return proj_yr_label, projection
