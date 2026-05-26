import io
import csv
import os
from datetime import datetime
from flask import Blueprint, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from app.parser.financial import parse_financial_workbooks
from app.parser.occupancy import parse_occupancy_report
from app.parser.ar_aging import parse_ar_aging_reports
from app.mapper.account_mapper import map_rows
from app.calculator.noi import calculate_noi
from app.calculator.economic_occ import enrich_eco_occ
from app.calculator.physical_occ import enrich_physical_occ
from app.exporter.main_workbook import build_main_workbook
from app.exporter.backup_workbook import build_backup_workbook
from app.exporter.validator import validate_both_workbooks
from app.storage.runs import new_run_id, save_run
from app.models import QualityCheck
from config import ECO_OCC_TARGET, QUARTERS, PROPERTY_NAME_MAP, MONTHS

bp = Blueprint("upload", __name__)
ALLOWED_EXT = {".xlsx", ".xls"}


@bp.route("/", methods=["GET"])
def index():
    return render_template("upload.html", eco_occ_target=ECO_OCC_TARGET * 100)


@bp.route("/", methods=["POST"])
def run_analysis():
    portfolio_name = request.form.get("portfolio_name", "Portfolio").strip() or "Portfolio"
    eco_occ_target      = float(request.form.get("eco_occ_target", ECO_OCC_TARGET * 100)) / 100
    use_budget_eco_occ  = request.form.get("use_budget_eco_occ") == "1"
    pm_names_raw        = request.form.get("pm_names", "").strip()
    excluded_raw   = request.form.get("excluded_properties", "").strip()
    carveout_raw   = request.form.get("carveout_properties", "").strip()

    excluded  = {p.strip().lower() for p in excluded_raw.splitlines() if p.strip()}
    carveouts = {p.strip().lower() for p in carveout_raw.splitlines() if p.strip()}

    fin_files = request.files.getlist("financial_files")
    occ_files = request.files.getlist("occupancy_file")

    if not fin_files or all(f.filename == "" for f in fin_files):
        flash("Please upload at least one financial statement workbook.")
        return redirect(url_for("upload.index"))

    os.makedirs("uploads", exist_ok=True)
    saved_paths = []
    pm_name_map = {}
    pm_lines = [l.strip() for l in pm_names_raw.splitlines() if l.strip()]

    for i, f in enumerate(fin_files):
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_EXT:
            continue
        fname = secure_filename(f.filename)
        path = os.path.join("uploads", fname)
        f.save(path)
        saved_paths.append(path)
        if i < len(pm_lines):
            pm_name_map[fname] = pm_lines[i]

    if not saved_paths:
        flash("No valid .xlsx files were uploaded.")
        return redirect(url_for("upload.index"))

    # Parse optional custom mapping CSV
    custom_mapping = None
    mapping_file = request.files.get("custom_mapping")
    if mapping_file and mapping_file.filename:
        content = mapping_file.read().decode("utf-8")
        reader = csv.DictReader(io.StringIO(content))
        custom_mapping = {}
        for row in reader:
            name = row.get("account_name", "").lower().strip()
            cat  = row.get("assigned_category", "")
            trt  = row.get("treatment", "")
            in_noi = row.get("include_in_noi", "").lower() == "yes"
            in_eco = row.get("include_in_eco_occ", "").lower() == "yes"
            if name and cat:
                custom_mapping[name] = (cat, trt, in_noi, in_eco)

    # ── Financial pipeline ────────────────────────────────────────────────────
    raw_rows, source_index = parse_financial_workbooks(saved_paths, pm_name_map)

    for _row in raw_rows:
        _row.property_name = PROPERTY_NAME_MAP.get(_row.property_name, _row.property_name)
    for _entry in source_index:
        _entry.property_name = PROPERTY_NAME_MAP.get(_entry.property_name, _entry.property_name)

    occ_rows = []
    for occ_file in occ_files:
        if occ_file and occ_file.filename:
            occ_path = os.path.join("uploads", secure_filename(occ_file.filename))
            occ_file.save(occ_path)
            occ_rows.extend(parse_occupancy_report(occ_path))

    for _row in occ_rows:
        _row.property_name = PROPERTY_NAME_MAP.get(_row.property_name, _row.property_name)

    # ── AR Aging pipeline ─────────────────────────────────────────────────────
    ar_rows = []
    ar_files = request.files.getlist("ar_aging_files")
    for ar_file in ar_files:
        if ar_file and ar_file.filename:
            ar_path = os.path.join("uploads", secure_filename(ar_file.filename))
            ar_file.save(ar_path)
            ar_rows.extend(parse_ar_aging_reports([ar_path]))

    # Apply PROPERTY_NAME_MAP to AR rows (parser defers to upload layer)
    for _row in ar_rows:
        _row.property_name = PROPERTY_NAME_MAP.get(_row.property_name, _row.property_name)

    # ── NOI / eco occ / physical occ ─────────────────────────────────────────
    mapped_rows, mapping_entries = map_rows(raw_rows, custom_mapping)
    kpis = calculate_noi(mapped_rows)
    kpis = enrich_eco_occ(mapped_rows, kpis)
    kpis = enrich_physical_occ(occ_rows, kpis)

    # Apply period filter
    period_filter = request.form.get("period_filter", "Full Year")
    selected_months_raw = request.form.getlist("selected_months")
    if period_filter in QUARTERS:
        allowed = set(QUARTERS[period_filter])
        kpis = [k for k in kpis if k.month in allowed]
    elif period_filter == "Selected Months" and selected_months_raw:
        allowed = {int(m) for m in selected_months_raw}
        kpis = [k for k in kpis if k.month in allowed]

    # Apply exclusions and carveouts
    kpis = [k for k in kpis if k.property_name.lower() not in excluded]
    for k in kpis:
        if k.property_name.lower() in carveouts:
            k.is_carveout = True

    for k in kpis:
        if k.eco_occ_pct is not None:
            if use_budget_eco_occ and k.budget_eco_occ_pct is not None:
                k.is_below_eco_occ_target = k.eco_occ_pct < k.budget_eco_occ_pct
            else:
                k.is_below_eco_occ_target = k.eco_occ_pct < eco_occ_target

    # ── Build workbooks ───────────────────────────────────────────────────────
    run_id  = new_run_id()
    run_dir = os.path.join("runs", run_id)
    os.makedirs(run_dir, exist_ok=True)

    safe_name   = "".join(c for c in portfolio_name if c.isalnum() or c in " _-").strip()
    main_path   = os.path.join(run_dir, f"{safe_name} Property Analysis.xlsx")
    backup_path = os.path.join(run_dir, f"{safe_name} Property Analysis backup.xlsx")

    build_main_workbook(kpis, portfolio_name, main_path, eco_occ_target,
                        ar_rows=ar_rows if ar_rows else None,
                        use_budget_eco_occ=use_budget_eco_occ)
    build_backup_workbook(mapped_rows, kpis, source_index, mapping_entries, [],
                          backup_path, eco_occ_target,
                          ar_rows=ar_rows if ar_rows else None)

    val_checks = validate_both_workbooks(main_path, backup_path)
    quality_checks = list(val_checks)

    years = sorted({k.year for k in kpis})
    props = sorted({k.property_name for k in kpis})
    pm_names_used = sorted({k.pm_name for k in kpis})

    # AR period metadata for history page
    ar_tr_periods  = sorted({(r.year, r.month) for r in ar_rows if r.receivable_type == "Tenant Rent"})
    ar_sub_periods = sorted({(r.year, r.month) for r in ar_rows if r.receivable_type == "Subsidy"})

    metadata = {
        "created_at": datetime.now().isoformat(),
        "portfolio_name": portfolio_name,
        "eco_occ_target": eco_occ_target,
        "use_budget_eco_occ": use_budget_eco_occ,
        "years": years,
        "properties": props,
        "num_properties": len(props),
        "pm_names": pm_names_used,
        "source_files": [os.path.basename(p) for p in saved_paths],
        "excluded_properties": list(excluded),
        "carveout_properties": list(carveouts),
        "main_workbook": os.path.basename(main_path),
        "backup_workbook": os.path.basename(backup_path),
        "ar_tenant_rent_periods": [f"{MONTHS[mo]}-{yr}" for (yr, mo) in ar_tr_periods],
        "ar_subsidy_periods":     [f"{MONTHS[mo]}-{yr}" for (yr, mo) in ar_sub_periods],
    }

    save_run(run_id, metadata, kpis, source_index, mapping_entries, quality_checks,
             ar_rows=ar_rows if ar_rows else None)

    # Clean up temp uploads
    for p in saved_paths:
        try:
            os.remove(p)
        except OSError:
            pass

    return redirect(url_for("results.show", run_id=run_id))
