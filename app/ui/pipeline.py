"""Analysis pipeline — Flask-free version of app/routes/upload.py business logic.

Accepts file bytes instead of Flask request.files objects. Returns run_id.
"""
import csv
import os
import re
import tempfile
from collections import defaultdict
from datetime import datetime

from app.parser.financial import parse_financial_workbooks, _infer_pm_from_filename
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
from config import (QUARTERS, PROPERTY_NAME_MAP, MONTHS,
                    PERMANENT_EXCLUSIONS, PROPERTY_METADATA, PROPERTY_PM_EXCLUSIONS)


def _save_bytes_to_temp(filename: str, data: bytes) -> str:
    """Write bytes to a NamedTemporaryFile and return the path."""
    suffix = os.path.splitext(filename)[1] or ".xlsx"
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path


def _clean_pm_name(filename: str) -> str:
    """Extract a short PM company name from a financial report filename.

    Strategy (in order):
    1. Known-PM lookup via _infer_pm_from_filename — matches any name in
       _KNOWN_PM_NAMES against the filename (case-insensitive substring).
       Handles cases like 'VOTP JSCO 12 month actual 2026.xlsx' → 'JSCO'
       where the stem starts with a property abbreviation.
    2. Regex fallback — strips common financial-report boilerplate suffixes
       ('12 month actual budget 2026', 'full year 2025', etc.) from the
       stem for any PM not yet in _KNOWN_PM_NAMES.

    Examples:
        'Solari 12 month actual budget 2026.xlsx'  → 'Solari'
        'ConAm 12 month actual budget 2026.xlsx'   → 'ConAm'
        'VOTP JSCO 12 month actual 2026.xlsx'      → 'JSCO'
        'SomeNewPM Full Year 2025.xlsx'            → 'SomeNewPM'
    """
    stem = (
        os.path.splitext(filename)[0]
        .replace("_", " ")
        .replace("-", " ")
        .strip()
    )
    # Step 1: known-PM lookup (returns the canonical name if found, else stem)
    inferred = _infer_pm_from_filename(filename)
    if inferred != stem:
        return inferred
    # Step 2: regex — strip boilerplate suffixes for unknown PM companies
    cleaned = re.sub(
        r'\s+(12\s+month|full[- ]year|full[- ]yr|annual|actual|budget|\d{4})\b.*$',
        "",
        stem,
        flags=re.IGNORECASE,
    ).strip()
    return cleaned if cleaned else stem


def _detect_partial_year(kpis) -> set[str]:
    """Auto-detect properties with fewer months than the max in their year."""
    months_by: dict = defaultdict(lambda: defaultdict(set))
    for k in kpis:
        if not k.is_carveout:
            months_by[k.property_name][k.year].add(k.month)
    all_years = {yr for pd in months_by.values() for yr in pd}
    max_per_year: dict[int, int] = {}
    for yr in all_years:
        counts = [len(months_by[p][yr]) for p in months_by if yr in months_by[p]]
        max_per_year[yr] = max(counts) if counts else 0
    partial: set[str] = set()
    for prop, yr_data in months_by.items():
        for yr, months in yr_data.items():
            if max_per_year.get(yr, 0) > 0 and len(months) < max_per_year[yr]:
                partial.add(prop)
    return partial


def run_analysis_pipeline(
    fin_files: list[tuple[str, bytes]],   # [(filename, bytes), ...]
    occ_files: list[tuple[str, bytes]],
    ar_files:  list[tuple[str, bytes]],
    settings:  dict,
) -> str:
    """Run the full analysis pipeline and return the saved run_id.

    settings keys:
        portfolio_name       str
        eco_occ_target       float  (e.g. 0.95)
        use_budget_eco_occ   bool
        pm_names             list[str]  (one per financial file, in order)
        excluded_properties  set[str]   (lowercase property names)
        carveout_properties  set[str]   (lowercase)
        stabilized_properties set[str]  (canonical names, exact match)
        period_filter        str  ("Full Year" | "Q1" | ... | "Selected Months")
        selected_months      list[int]
        custom_mapping       dict | None
    """
    ALLOWED_EXT = {".xlsx", ".xls"}
    portfolio_name      = settings.get("portfolio_name", "Portfolio").strip() or "Portfolio"
    eco_occ_target      = float(settings.get("eco_occ_target", 0.95))
    use_budget_eco_occ  = bool(settings.get("use_budget_eco_occ", False))
    pm_names            = list(settings.get("pm_names", []))
    excluded            = set(settings.get("excluded_properties", set())) | PERMANENT_EXCLUSIONS
    carveouts           = set(settings.get("carveout_properties", set()))
    manual_stabilized   = set(settings.get("stabilized_properties", set()))
    period_filter       = settings.get("period_filter", "Full Year")
    selected_months     = list(settings.get("selected_months", []))
    custom_mapping      = settings.get("custom_mapping")

    # ── Save financial files to temp paths ───────────────────────────────────
    saved_paths: list[str] = []
    pm_name_map: dict[str, str] = {}
    for i, (fname, data) in enumerate(fin_files):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in ALLOWED_EXT:
            continue
        path = _save_bytes_to_temp(fname, data)
        saved_paths.append(path)
        if i < len(pm_names) and pm_names[i]:
            pm_name_map[os.path.basename(path)] = pm_names[i]
        else:
            # Derive a short PM name from the original filename (e.g. "Solari"
            # from "Solari 12 month actual budget 2026.xlsx") so the PM column
            # shows a concise company name rather than the full filename stem.
            pm_name_map[os.path.basename(path)] = _clean_pm_name(fname)

    if not saved_paths:
        raise ValueError("No valid .xlsx files were provided.")

    occ_paths: list[str] = []
    ar_paths:  list[str] = []
    try:
        # ── Save occupancy and AR files ───────────────────────────────────────
        occ_paths = [_save_bytes_to_temp(fn, d) for fn, d in occ_files if fn]

        # Build AR paths AND a map of temp_path → original_filename so the
        # AR aging parser can extract PM name + receivable type from the real
        # filename (e.g. "Solari_AR Aging_Subsidy_03_2026.xlsx") rather than
        # from the random temp basename (e.g. "tmpabcde123.xlsx").
        ar_paths: list[str] = []
        ar_original_names: dict[str, str] = {}
        for fn, d in ar_files:
            if not fn:
                continue
            path = _save_bytes_to_temp(fn, d)
            ar_paths.append(path)
            ar_original_names[path] = fn

        # ── Parse ────────────────────────────────────────────────────────────
        raw_rows, source_index = parse_financial_workbooks(saved_paths, pm_name_map)

        for _row in raw_rows:
            _row.property_name = PROPERTY_NAME_MAP.get(_row.property_name, _row.property_name)
        for _entry in source_index:
            _entry.property_name = PROPERTY_NAME_MAP.get(_entry.property_name, _entry.property_name)

        if PROPERTY_PM_EXCLUSIONS:
            raw_rows = [
                r for r in raw_rows
                if PROPERTY_PM_EXCLUSIONS.get(
                    r.property_name.lower(), r.pm_name
                ).lower() == r.pm_name.lower()
            ]

        occ_rows = []
        for path in occ_paths:
            rows = parse_occupancy_report(path)
            for r in rows:
                r.property_name = PROPERTY_NAME_MAP.get(r.property_name, r.property_name)
            occ_rows.extend(rows)

        ar_rows = []
        if ar_paths:
            rows = parse_ar_aging_reports(ar_paths, original_names=ar_original_names)
            for r in rows:
                r.property_name = PROPERTY_NAME_MAP.get(r.property_name, r.property_name)
            ar_rows.extend(rows)

        # ── Calculate ──────────────────────────────────────────────────────────
        mapped_rows, mapping_entries = map_rows(raw_rows, custom_mapping)
        kpis = calculate_noi(mapped_rows)
        kpis = enrich_eco_occ(mapped_rows, kpis)
        kpis = enrich_physical_occ(occ_rows, kpis)

        # Period filter
        if period_filter in QUARTERS:
            allowed = set(QUARTERS[period_filter])
            kpis = [k for k in kpis if k.month in allowed]
        elif period_filter == "Selected Months" and selected_months:
            allowed = set(selected_months)
            kpis = [k for k in kpis if k.month in allowed]

        # Exclusions and carveouts
        kpis    = [k for k in kpis    if k.property_name.lower() not in excluded]
        ar_rows = [r for r in ar_rows if r.property_name.lower() not in excluded]
        for k in kpis:
            if k.property_name.lower() in carveouts:
                k.is_carveout = True

        # Property metadata
        for k in kpis:
            meta = PROPERTY_METADATA.get(k.property_name, {})
            k.city         = meta.get("city", "")
            k.tenancy_type = meta.get("tenancy_type", "")

        # Partial-year detection
        auto_partial = _detect_partial_year(kpis)
        partial_year_props = auto_partial | manual_stabilized
        for k in kpis:
            if k.property_name in partial_year_props:
                k.is_partial_year = True

        for k in kpis:
            if k.eco_occ_pct is not None:
                if use_budget_eco_occ and k.budget_eco_occ_pct is not None:
                    k.is_below_eco_occ_target = k.eco_occ_pct < k.budget_eco_occ_pct
                else:
                    k.is_below_eco_occ_target = k.eco_occ_pct < eco_occ_target

        # ── Build workbooks ────────────────────────────────────────────────────
        run_id  = new_run_id()
        run_dir = os.path.join("runs", run_id)
        os.makedirs(run_dir, exist_ok=True)

        safe_name   = re.sub(r"[^\w\s\-]", "", portfolio_name).strip()
        main_path   = os.path.join(run_dir, f"{safe_name} Property Analysis.xlsx")
        backup_path = os.path.join(run_dir, f"{safe_name} Property Analysis backup.xlsx")

        build_main_workbook(kpis, portfolio_name, main_path, eco_occ_target,
                            ar_rows=ar_rows or None,
                            use_budget_eco_occ=use_budget_eco_occ)
        build_backup_workbook(mapped_rows, kpis, source_index, mapping_entries, [],
                              backup_path, eco_occ_target,
                              ar_rows=ar_rows or None)

        val_checks = validate_both_workbooks(main_path, backup_path)
        quality_checks = list(val_checks)

        years         = sorted({k.year for k in kpis})
        props         = sorted({k.property_name for k in kpis})
        pm_names_used = sorted({k.pm_name for k in kpis})

        ar_tr_periods  = sorted({(r.year, r.month) for r in ar_rows
                                  if r.receivable_type == "Tenant Rent"})
        ar_sub_periods = sorted({(r.year, r.month) for r in ar_rows
                                  if r.receivable_type == "Subsidy"})

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
            "excluded_properties":     sorted(excluded - PERMANENT_EXCLUSIONS),
            "carveout_properties":     sorted(carveouts),
            "partial_year_properties": sorted(partial_year_props),
            "manually_stabilized":     sorted(manual_stabilized),
            "auto_detected_partial":   sorted(auto_partial),
            "main_workbook":           os.path.basename(main_path),
            "backup_workbook":         os.path.basename(backup_path),
            "ar_tenant_rent_periods":  [f"{MONTHS[mo]}-{yr}" for (yr, mo) in ar_tr_periods],
            "ar_subsidy_periods":      [f"{MONTHS[mo]}-{yr}" for (yr, mo) in ar_sub_periods],
        }

        save_run(run_id, metadata, kpis, source_index, mapping_entries,
                 quality_checks, ar_rows=ar_rows or None)

        return run_id

    finally:
        # Clean up all temp files
        for p in saved_paths + occ_paths + ar_paths:
            try:
                os.remove(p)
            except OSError:
                pass
