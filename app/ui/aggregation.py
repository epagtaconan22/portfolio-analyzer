"""KPI and AR aggregation helpers — extracted from routes/results.py."""

from typing import Optional


# ── Quarter helpers ────────────────────────────────────────────────────────────

def month_to_quarter(month: int) -> int:
    return (month - 1) // 3 + 1


def quarter_label(year: int, quarter: int) -> str:
    return f"Q{quarter} - {year}"


def quarter_months(quarter: int) -> set:
    base = (quarter - 1) * 3 + 1
    return {base, base + 1, base + 2}


# ── AR Aging helpers ───────────────────────────────────────────────────────────

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


def ar_period_label(year: int, month: int) -> str:
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


def agg_ar(ar_rows: list, receivable_type: str,
           year: int, month: int) -> Optional[dict]:
    rows = [r for r in ar_rows
            if r["receivable_type"] == receivable_type
            and r["year"] == year and r["month"] == month]
    if not rows:
        return None
    charge_amount = sum(r["charge_amount"] for r in rows)
    total_over_60 = sum(r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed":   sum(r["current_owed"] for r in rows),
        "prepayments":    sum(r["prepayments"]  for r in rows),
        "pct_overdue":    (total_over_60 / charge_amount) if charge_amount > 0 else 0.0,
        "property_count": len({r["property_name"] for r in rows}),
    }


def agg_ar_for_prop(ar_rows: list, property_name: str,
                    receivable_type: str, year: int, month: int) -> Optional[dict]:
    rows = [r for r in ar_rows
            if r["property_name"] == property_name
            and r["receivable_type"] == receivable_type
            and r["year"] == year and r["month"] == month]
    if not rows:
        return None
    charge  = sum(r["charge_amount"] for r in rows)
    over_60 = sum(r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed": sum(r["current_owed"] for r in rows),
        "prepayments":  sum(r["prepayments"]  for r in rows),
        "pct_overdue":  (over_60 / charge) if charge and charge > 0 else 0.0,
    }


def ar_yoy_delta(curr: dict, prev: dict) -> dict:
    pct_delta = None
    if curr.get("pct_overdue") is not None and prev.get("pct_overdue") is not None:
        pct_delta = curr["pct_overdue"] - prev["pct_overdue"]
    return {
        "current_owed_delta": curr["current_owed"] - prev["current_owed"],
        "prepayments_delta":  curr["prepayments"]  - prev["prepayments"],
        "pct_overdue_delta":  pct_delta,
    }


def pct_delta(curr: Optional[dict], prev: Optional[dict]) -> Optional[float]:
    if curr and prev:
        c = curr.get("pct_overdue")
        p = prev.get("pct_overdue")
        if c is not None and p is not None:
            return c - p
    return None


# ── KPI aggregation ────────────────────────────────────────────────────────────

def agg_kpis(kpi_dicts: list) -> dict:
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

    actual_noi = (actual_income - actual_expenses
                  if actual_income is not None and actual_expenses is not None else None)
    budget_noi = (budget_income - budget_expenses
                  if budget_income is not None and budget_expenses is not None else None)

    net_coll = (gpr - (vacancy or 0) - (concessions or 0) - (bad_debt or 0)
                if gpr is not None else None)
    eco_occ  = (net_coll / gpr) if (net_coll is not None and gpr) else None

    bud_eco_vals = [k["budget_eco_occ_pct"] for k in kpi_dicts
                    if k.get("budget_eco_occ_pct") is not None]
    bud_eco      = sum(bud_eco_vals) / len(bud_eco_vals) if bud_eco_vals else None
    eco_occ_var  = (eco_occ - bud_eco
                    if eco_occ is not None and bud_eco is not None else None)

    noi_var     = (actual_noi - budget_noi
                   if actual_noi is not None and budget_noi is not None else None)
    noi_var_pct = (noi_var / abs(budget_noi)
                   if noi_var is not None and budget_noi else None)

    _paired = [(k["occupied_units"], k["total_units"])
               for k in kpi_dicts
               if k.get("occupied_units") is not None and k.get("total_units") is not None]
    phys_occ = (sum(p[0] for p in _paired) / sum(p[1] for p in _paired)
                if _paired and sum(p[1] for p in _paired) > 0 else None)

    total_unit_months = sum(k["total_units"] for k in kpi_dicts
                            if k.get("total_units") is not None) or None

    income_pu  = (actual_income   / total_unit_months
                  if actual_income   is not None and total_unit_months else None)
    expense_pu = (actual_expenses / total_unit_months
                  if actual_expenses is not None and total_unit_months else None)
    noi_pu     = (actual_noi      / total_unit_months
                  if actual_noi      is not None and total_unit_months else None)

    def _safe_pct(num, denom):
        if num is None or denom is None or denom == 0:
            return None
        return num / abs(denom)

    inc_var = ((actual_income - budget_income)
               if actual_income is not None and budget_income is not None else None)
    exp_var = ((actual_expenses - budget_expenses)
               if actual_expenses is not None and budget_expenses is not None else None)

    total_units = next((k["total_units"] for k in kpi_dicts
                        if k.get("total_units") is not None), None)

    return dict(
        actual_income=actual_income,
        budget_income=budget_income,
        income_variance=inc_var,
        income_variance_pct=_safe_pct(inc_var, budget_income),
        actual_expenses=actual_expenses,
        budget_expenses=budget_expenses,
        expense_variance=exp_var,
        expense_variance_pct=_safe_pct(exp_var, budget_expenses),
        actual_noi=actual_noi,
        budget_noi=budget_noi,
        noi_variance=noi_var,
        noi_variance_pct=noi_var_pct,
        eco_occ_pct=eco_occ,
        budget_eco_occ_pct=bud_eco,
        eco_occ_variance=eco_occ_var,
        physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ
                     if phys_occ is not None and eco_occ is not None else None),
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
