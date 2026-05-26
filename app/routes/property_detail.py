from flask import Blueprint, render_template, abort
from app.storage.runs import load_run

bp = Blueprint("property_detail", __name__)

_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def _ar_period_label(year: int, month: int) -> str:
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


@bp.route("/results/<run_id>/property/<property_name>")
def show(run_id, property_name):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    prop_kpis = [k for k in data["kpis"] if k["property_name"] == property_name]
    if not prop_kpis:
        abort(404)

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

    prop_ar.sort(key=lambda r: (r["receivable_type"], r["year"], r["month"]))

    # Split into sub-dicts by receivable type for template simplicity
    ar_by_type = {}
    for r in prop_ar:
        ar_by_type.setdefault(r["receivable_type"], []).append(r)

    return render_template(
        "property_detail.html",
        run_id=run_id,
        property_name=property_name,
        kpis=prop_kpis,
        meta=data["metadata"],
        ar_by_type=ar_by_type,
    )
