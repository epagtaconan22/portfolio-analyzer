from flask import Blueprint, render_template, abort
from app.storage.runs import load_run

bp = Blueprint("property_detail", __name__)


@bp.route("/results/<run_id>/property/<property_name>")
def show(run_id, property_name):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    prop_kpis = [k for k in data["kpis"] if k["property_name"] == property_name]
    if not prop_kpis:
        abort(404)

    return render_template(
        "property_detail.html",
        run_id=run_id,
        property_name=property_name,
        kpis=prop_kpis,
        meta=data["metadata"],
    )
