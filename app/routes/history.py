from flask import Blueprint, render_template, redirect, url_for
from app.storage.runs import list_runs, delete_run

bp = Blueprint("history", __name__)


@bp.route("/history")
def index():
    runs = list_runs()
    return render_template("history.html", runs=runs)


@bp.route("/history/<run_id>/delete", methods=["POST"])
def delete(run_id):
    delete_run(run_id)
    return redirect(url_for("history.index"))
