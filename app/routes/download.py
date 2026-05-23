import io
import os
import zipfile
from flask import Blueprint, send_file, abort
from app.storage.runs import load_run

bp = Blueprint("download", __name__)


@bp.route("/results/<run_id>/download")
def download(run_id):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    meta      = data["metadata"]
    run_dir   = os.path.join("runs", run_id)
    main_wb   = os.path.join(run_dir, meta.get("main_workbook", ""))
    backup_wb = os.path.join(run_dir, meta.get("backup_workbook", ""))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.isfile(main_wb):
            zf.write(main_wb, os.path.basename(main_wb))
        if os.path.isfile(backup_wb):
            zf.write(backup_wb, os.path.basename(backup_wb))
    buf.seek(0)

    portfolio_name = meta.get("portfolio_name", "Portfolio")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{portfolio_name} Analysis Workbooks.zip",
        mimetype="application/zip",
    )
