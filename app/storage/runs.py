"""Save, load, list, and delete analysis runs from the runs/ directory."""

import json
import os
import shutil
import uuid
from datetime import datetime
from dataclasses import asdict
from app.models import PropertyPeriodKPIs, SourceIndexEntry, MappingEntry, QualityCheck

RUNS_DIR = "runs"

def new_run_id() -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{ts}_{uuid.uuid4().hex[:6]}"

def save_run(
    run_id: str,
    metadata: dict,
    kpis: list[PropertyPeriodKPIs],
    source_index: list[SourceIndexEntry],
    mapping_entries: list[MappingEntry],
    quality_checks: list[QualityCheck],
) -> str:
    """Saves run data to runs/<run_id>/. Returns the run directory path."""
    run_dir = os.path.join(RUNS_DIR, run_id)
    os.makedirs(run_dir, exist_ok=True)

    with open(os.path.join(run_dir, "metadata.json"), "w") as f:
        json.dump(metadata, f, indent=2, default=str)

    with open(os.path.join(run_dir, "kpis.json"), "w") as f:
        json.dump([asdict(k) for k in kpis], f, indent=2, default=str)

    with open(os.path.join(run_dir, "source_index.json"), "w") as f:
        json.dump([asdict(e) for e in source_index], f, indent=2, default=str)

    with open(os.path.join(run_dir, "mapping_entries.json"), "w") as f:
        json.dump([asdict(e) for e in mapping_entries], f, indent=2, default=str)

    with open(os.path.join(run_dir, "quality_checks.json"), "w") as f:
        json.dump([asdict(c) for c in quality_checks], f, indent=2, default=str)

    return run_dir


def load_run(run_id: str) -> dict:
    """Returns a dict with keys: metadata, kpis, source_index, mapping_entries, quality_checks."""
    run_dir = os.path.join(RUNS_DIR, run_id)
    result = {}
    for key in ("metadata", "kpis", "source_index", "mapping_entries", "quality_checks"):
        path = os.path.join(run_dir, f"{key}.json")
        with open(path) as f:
            result[key] = json.load(f)
    return result


def list_runs() -> list[dict]:
    """Returns list of metadata dicts for all runs, sorted newest first."""
    if not os.path.isdir(RUNS_DIR):
        return []
    runs = []
    for run_id in os.listdir(RUNS_DIR):
        meta_path = os.path.join(RUNS_DIR, run_id, "metadata.json")
        if os.path.isfile(meta_path):
            with open(meta_path) as f:
                meta = json.load(f)
            meta["run_id"] = run_id
            runs.append(meta)
    return sorted(runs, key=lambda r: r.get("created_at", ""), reverse=True)


def delete_run(run_id: str) -> None:
    run_dir = os.path.join(RUNS_DIR, run_id)
    if os.path.isdir(run_dir):
        shutil.rmtree(run_dir)
