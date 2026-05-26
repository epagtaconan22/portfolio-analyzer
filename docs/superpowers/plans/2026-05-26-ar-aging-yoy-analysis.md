# AR Aging YoY Analysis — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add AR Aging YoY trend analysis to the portfolio-analyzer app: parser for Yardi AR Aging exports, pipeline integration, web dashboard section, property detail section, and new tabs in both Excel workbooks.

**Architecture:** Additive changes only — new `ARAgingRow` dataclass, new `app/parser/ar_aging.py`, and `ar_rows=None` optional parameter threaded through existing `save_run`, `build_main_workbook`, and `build_backup_workbook`. Dashboard and property detail templates get new Jinja2 sections guarded by `{% if ar_summary %}`. The validator tab-set constants are updated so the AR Aging / AR_Aging_Detail tabs are always expected.

**Tech Stack:** Python 3.11+, openpyxl 3.x, Flask/Jinja2, pytest. No new dependencies.

---

## File Map

| File | Action | Description |
|---|---|---|
| `app/models.py` | Modify | Add `ARAgingRow` dataclass |
| `app/parser/ar_aging.py` | **Create** | Yardi AR Aging file parser |
| `tests/test_ar_aging_parser.py` | **Create** | Parser unit tests (TDD) |
| `app/storage/runs.py` | Modify | Add `ar_rows` param; save/load `ar_aging.json` |
| `app/templates/upload.html` | Modify | Add AR Aging drop zone section |
| `app/routes/upload.py` | Modify | Parse AR files; pass to exporters and `save_run` |
| `app/routes/results.py` | Modify | Add `_agg_ar`, `_agg_ar_for_prop`, `_ar_yoy_delta`, `_ar_period_label`; pass AR context to template |
| `app/templates/dashboard.html` | Modify | Add AR Aging portfolio summary + property table section |
| `app/routes/property_detail.py` | Modify | Pass `prop_ar_rows` to template |
| `app/templates/property_detail.html` | Modify | Add AR Aging detail section |
| `app/exporter/main_workbook.py` | Modify | Add `ar_rows=None` param; new `_build_ar_aging` function; new "AR Aging" tab |
| `app/exporter/backup_workbook.py` | Modify | Add `ar_rows=None` param; new `_build_ar_aging_detail` function; new "AR_Aging_Detail" tab |
| `app/exporter/validator.py` | Modify | Add "AR Aging" to `_MAIN_TABS`; add "AR_Aging_Detail" to `_BACKUP_TABS` |

---

## Reference: Yardi AR Aging File Format

Sheet name: `Report1`

| Row index (0-based) | Content |
|---|---|
| 0 | Report title |
| 1 | Property line |
| 2 | `"Post To(MM/YY): 03/2024"` |
| 3 | Column header row (top half) |
| 4 | Column header row (bottom half) |
| 5… | Data rows — stop at first `None` in col A or row starting with `"Grand Total"` |

Column mapping (0-indexed within each data row):

| Index | Field |
|---|---|
| 0 | Property Name with code suffix: `"Alora Family (alora)"` |
| 1 | Charge Amount |
| 2 | Current Owed |
| 3 | 0-30 Owed |
| 4 | 31-60 Owed |
| 5 | 61-90 Owed |
| 6 | Over 90 Owed |
| 7 | Pre-payments (credits — always ≤ 0) |
| 8 | Suspense (always 0, ignored) |

---

## Task 1: ARAgingRow Model

**Files:**
- Modify: `app/models.py`

- [ ] **Step 1: Append `ARAgingRow` to `app/models.py`**

Add the following after the existing `QualityCheck` dataclass (at the end of the file):

```python
@dataclass
class ARAgingRow:
    property_name: str       # After PROPERTY_NAME_MAP normalization
    pm_name: str             # Extracted from filename prefix before first "_"
    source_file: str         # Basename of source file
    receivable_type: str     # "Tenant Rent" | "Subsidy"
    year: int
    month: int               # 1–12
    charge_amount: float     # Col 1
    current_owed: float      # Col 2
    owed_0_30: float         # Col 3
    owed_31_60: float        # Col 4
    owed_61_90: float        # Col 5
    owed_over_90: float      # Col 6
    prepayments: float       # Col 7 (negative = credits)

    @property
    def total_overdue(self) -> float:
        """31-60 + 61-90 + Over-90 — amounts past the 0-30 bucket."""
        return self.owed_31_60 + self.owed_61_90 + self.owed_over_90

    @property
    def pct_overdue(self) -> Optional[float]:
        """% of charge_amount that is >30 days past due."""
        if self.charge_amount and self.charge_amount > 0:
            return self.total_overdue / self.charge_amount
        return None
```

- [ ] **Step 2: Verify import**

```powershell
cd C:\Users\erwin\Desktop\portfolio-analyzer
python -c "from app.models import ARAgingRow; r = ARAgingRow('Prop','PM','f.xlsx','Tenant Rent',2024,3,10000,8000,500,300,200,100,-50); print(r.total_overdue, r.pct_overdue)"
```

Expected output: `600 0.06`

- [ ] **Step 3: Commit**

```powershell
git add app/models.py
git commit -m "feat: add ARAgingRow dataclass"
```

---

## Task 2: AR Aging Parser (TDD)

**Files:**
- Create: `app/parser/ar_aging.py`
- Create: `tests/test_ar_aging_parser.py`

- [ ] **Step 1: Write failing tests — create `tests/test_ar_aging_parser.py`**

```python
"""Tests for the Yardi AR Aging file parser."""

import pytest
import openpyxl

from app.models import ARAgingRow
from app.parser.ar_aging import parse_ar_aging_reports


# ── Fixtures ─────────────────────────────────────────────────────────────────

def _make_ar_wb(tmp_path, filename, rows):
    """Helper: create a synthetic Yardi AR Aging workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report1"
    ws.append(["Affordable Aging Detail"])                              # row 0
    ws.append(["Property: Affirmed Property List (affirmed)"])          # row 1
    period = f"Post To(MM/YY): {filename.split('_')[-2].zfill(2)}/{filename.split('_')[-1].split('.')[0]}"
    ws.append([period])                                                  # row 2
    ws.append(["Property Name", "Charge Amount", "Current Owed",       # row 3
               "0-30 Owed", "31-60 Owed", "61-90 Owed",
               "Over 90 Owed", "Pre-payments", "Suspense"])
    ws.append([""] * 9)                                                  # row 4
    for r in rows:
        ws.append(r)
    ws.append([None] * 9)                                               # blank
    ws.append(["Grand Total"] + [0] * 8)                               # grand total
    path = tmp_path / filename
    wb.save(str(path))
    return str(path)


@pytest.fixture
def tenant_rent_wb(tmp_path):
    return _make_ar_wb(tmp_path, "Solari_AR Aging_Tenant Rent_03_2024.xlsx", [
        ["Alora Family (alora)",  10000, 8000, 500, 300, 200, 100, -50,   0],
        ["Beechwood (beech)",     20000, 15000, 1000, 800, 600, 400, -100, 0],
    ])


@pytest.fixture
def subsidy_wb(tmp_path):
    return _make_ar_wb(tmp_path, "Solari_AR Aging_Subsidy_03_2024.xlsx", [
        ["Alora Family (alora)", 5000, 4000, 200, 100, 50, 25, 0, 0],
    ])


@pytest.fixture
def tenant_receivable_wb(tmp_path):
    """Type normalization: 'Tenant Receivable' → 'Tenant Rent'."""
    return _make_ar_wb(tmp_path, "ConAm_AR Aging_Tenant Receivable_06_2024.xlsx", [
        ["Alora Family (alora)", 10000, 8000, 500, 300, 200, 100, -50, 0],
    ])


@pytest.fixture
def subsidy_receivable_wb(tmp_path):
    """Type normalization: 'Subsidy Receivable' → 'Subsidy'."""
    return _make_ar_wb(tmp_path, "ConAm_AR Aging_Subsidy Receivable_06_2024.xlsx", [
        ["Alora Family (alora)", 5000, 4000, 200, 100, 50, 25, 0, 0],
    ])


# ── Tests ────────────────────────────────────────────────────────────────────

def test_returns_ar_aging_row_instances(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert len(rows) == 2
    assert all(isinstance(r, ARAgingRow) for r in rows)


def test_receivable_type_tenant_rent(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.receivable_type == "Tenant Rent" for r in rows)


def test_receivable_type_subsidy(subsidy_wb):
    rows = parse_ar_aging_reports([subsidy_wb])
    assert all(r.receivable_type == "Subsidy" for r in rows)


def test_type_normalization_tenant_receivable(tenant_receivable_wb):
    rows = parse_ar_aging_reports([tenant_receivable_wb])
    assert all(r.receivable_type == "Tenant Rent" for r in rows)


def test_type_normalization_subsidy_receivable(subsidy_receivable_wb):
    rows = parse_ar_aging_reports([subsidy_receivable_wb])
    assert all(r.receivable_type == "Subsidy" for r in rows)


def test_pm_name_from_filename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.pm_name == "Solari" for r in rows)


def test_year_month_from_filename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.year == 2024 and r.month == 3 for r in rows)


def test_property_code_suffix_stripped(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    prop_names = {r.property_name for r in rows}
    assert "Alora Family" in prop_names
    assert "Beechwood" in prop_names
    assert not any("(" in name for name in prop_names)


def test_numeric_fields(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    assert alora.charge_amount  == pytest.approx(10000)
    assert alora.current_owed   == pytest.approx(8000)
    assert alora.owed_0_30      == pytest.approx(500)
    assert alora.owed_31_60     == pytest.approx(300)
    assert alora.owed_61_90     == pytest.approx(200)
    assert alora.owed_over_90   == pytest.approx(100)
    assert alora.prepayments    == pytest.approx(-50)


def test_grand_total_row_excluded(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert not any("Grand Total" in r.property_name for r in rows)


def test_blank_row_stops_iteration(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    # blank row (None in col A) terminates parsing — only data rows returned
    assert all(r.property_name for r in rows)


def test_total_overdue_computed(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    assert alora.total_overdue == pytest.approx(600)   # 300+200+100


def test_pct_overdue_computed(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    alora = next(r for r in rows if "Alora" in r.property_name)
    assert alora.pct_overdue == pytest.approx(0.06)    # 600/10000


def test_none_numeric_treated_as_zero(tmp_path):
    """None values in numeric columns become 0.0 (not crash)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report1"
    ws.append(["Affordable Aging Detail"])
    ws.append(["Property: Affirmed Property List (affirmed)"])
    ws.append(["Post To(MM/YY): 03/2024"])
    ws.append(["Property Name", "Charge Amount", "Current Owed",
               "0-30 Owed", "31-60 Owed", "61-90 Owed",
               "Over 90 Owed", "Pre-payments", "Suspense"])
    ws.append([""] * 9)
    ws.append(["Test Prop (test)", 5000, 5000, None, None, None, None, None, 0])
    ws.append([None] * 9)
    ws.append(["Grand Total"] + [0] * 8)
    path = tmp_path / "PM_AR Aging_Subsidy_03_2024.xlsx"
    wb.save(str(path))
    rows = parse_ar_aging_reports([str(path)])
    assert len(rows) == 1
    assert rows[0].owed_31_60   == 0.0
    assert rows[0].prepayments  == 0.0
    assert rows[0].pct_overdue  == 0.0   # 0/5000


def test_source_file_is_basename(tenant_rent_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb])
    assert all(r.source_file == "Solari_AR Aging_Tenant Rent_03_2024.xlsx" for r in rows)


def test_multiple_files_combined(tenant_rent_wb, subsidy_wb):
    rows = parse_ar_aging_reports([tenant_rent_wb, subsidy_wb])
    types = {r.receivable_type for r in rows}
    assert types == {"Tenant Rent", "Subsidy"}
```

- [ ] **Step 2: Run tests to confirm they fail**

```powershell
cd C:\Users\erwin\Desktop\portfolio-analyzer
pytest tests/test_ar_aging_parser.py -v
```

Expected: All FAIL with `ModuleNotFoundError: No module named 'app.parser.ar_aging'`

- [ ] **Step 3: Create `app/parser/ar_aging.py`**

```python
"""Parses Yardi AR Aging export files (Tenant Rent and Subsidy) into ARAgingRow records."""

import os
import re
from typing import Optional
import openpyxl

from app.models import ARAgingRow
from config import PROPERTY_NAME_MAP, MONTHS

# Map lowercase type strings from filename to canonical values
_TYPE_NORMALIZE: dict[str, str] = {
    "tenant rent":        "Tenant Rent",
    "tenant receivable":  "Tenant Rent",
    "subsidy":            "Subsidy",
    "subsidy receivable": "Subsidy",
}


def parse_ar_aging_reports(file_paths: list[str]) -> list[ARAgingRow]:
    """Parse one or more Yardi AR Aging export files. Returns combined list of ARAgingRow."""
    results: list[ARAgingRow] = []
    for path in file_paths:
        results.extend(_parse_one(path))
    return results


# ── Per-file parsing ──────────────────────────────────────────────────────────

def _parse_one(path: str) -> list[ARAgingRow]:
    fname = os.path.basename(path)
    stem  = os.path.splitext(fname)[0]

    pm_name, receivable_type, year, month = _parse_filename(stem)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Report1"] if "Report1" in wb.sheetnames else wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    # Fallback: read period from sheet row 3 if filename didn't match
    if year is None or month is None:
        year, month = _parse_period_from_sheet(all_rows)
    if receivable_type is None:
        receivable_type = _infer_type_from_stem(stem)

    rows: list[ARAgingRow] = []
    # Data rows start at index 5 (0-based) — after 5 header rows
    for raw_row in all_rows[5:]:
        col_a = raw_row[0]
        if col_a is None:
            break
        raw_name = str(col_a).strip()
        if not raw_name or raw_name.startswith("Grand Total"):
            break

        # Strip property code suffix "(code)" from property name
        property_name = re.sub(r'\s*\([^)]+\)\s*$', '', raw_name).strip()
        if not property_name:
            continue

        # Apply PROPERTY_NAME_MAP normalization
        property_name = PROPERTY_NAME_MAP.get(property_name, property_name)

        rows.append(ARAgingRow(
            property_name=property_name,
            pm_name=pm_name or "",
            source_file=fname,
            receivable_type=receivable_type or "Unknown",
            year=year or 0,
            month=month or 0,
            charge_amount=_to_float(raw_row[1]),
            current_owed=_to_float(raw_row[2]),
            owed_0_30=_to_float(raw_row[3]),
            owed_31_60=_to_float(raw_row[4]),
            owed_61_90=_to_float(raw_row[5]),
            owed_over_90=_to_float(raw_row[6]),
            prepayments=_to_float(raw_row[7]),
        ))

    return rows


# ── Filename parsing ──────────────────────────────────────────────────────────

def _parse_filename(stem: str) -> tuple[Optional[str], Optional[str], Optional[int], Optional[int]]:
    """
    Parse stem like "Solari_AR Aging_Tenant Rent_03_2024".
    Returns (pm_name, receivable_type, year, month).
    Returns (None, None, None, None) if the pattern doesn't match.
    """
    parts = stem.split("_")
    # Minimum parts: [pm, "AR Aging", type_segment, month_str, year_str]
    if len(parts) < 5:
        return None, None, None, None

    try:
        month = int(parts[-2])
        year  = int(parts[-1])
    except ValueError:
        return None, None, None, None

    if not (1 <= month <= 12 and 2000 <= year <= 2100):
        return None, None, None, None

    pm_name  = parts[0]
    type_raw = " ".join(parts[2:-2]).lower().strip()
    receivable_type = _TYPE_NORMALIZE.get(type_raw)  # None if unrecognized

    return pm_name, receivable_type, year, month


def _parse_period_from_sheet(all_rows: list) -> tuple[Optional[int], Optional[int]]:
    """Fallback: parse period from row index 2 — 'Post To(MM/YY): 03/2024'."""
    if len(all_rows) < 3 or all_rows[2][0] is None:
        return None, None
    m = re.search(r'(\d{1,2})/(\d{4})', str(all_rows[2][0]))
    if m:
        return int(m.group(2)), int(m.group(1))   # (year, month)
    return None, None


def _infer_type_from_stem(stem: str) -> str:
    """Infer receivable type from filename stem keywords when pattern doesn't match."""
    stem_lower = stem.lower()
    if "subsidy" in stem_lower:
        return "Subsidy"
    return "Tenant Rent"


# ── Utilities ─────────────────────────────────────────────────────────────────

def _to_float(val) -> float:
    """Convert a cell value to float; treat None as 0.0."""
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
```

- [ ] **Step 4: Run tests to confirm they all pass**

```powershell
pytest tests/test_ar_aging_parser.py -v
```

Expected: All PASS (15+ tests green).

- [ ] **Step 5: Commit**

```powershell
git add app/parser/ar_aging.py tests/test_ar_aging_parser.py
git commit -m "feat: AR aging parser with TDD (Yardi AR Aging export format)"
```

---

## Task 3: Storage Update

**Files:**
- Modify: `app/storage/runs.py`

- [ ] **Step 1: Update `app/storage/runs.py`**

Replace the entire file with the following (adds `ar_rows` parameter to `save_run` and `"ar_aging"` key to `load_run`):

```python
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
    ar_rows: list | None = None,
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

    if ar_rows:
        with open(os.path.join(run_dir, "ar_aging.json"), "w") as f:
            json.dump([asdict(r) for r in ar_rows], f, indent=2, default=str)

    return run_dir


def load_run(run_id: str) -> dict:
    """Returns a dict with keys: metadata, kpis, source_index, mapping_entries,
    quality_checks, ar_aging (empty list if no AR data was uploaded)."""
    run_dir = os.path.join(RUNS_DIR, run_id)
    result = {}
    for key in ("metadata", "kpis", "source_index", "mapping_entries", "quality_checks"):
        path = os.path.join(run_dir, f"{key}.json")
        with open(path) as f:
            result[key] = json.load(f)

    # AR aging is optional — older runs won't have this file
    ar_path = os.path.join(run_dir, "ar_aging.json")
    if os.path.isfile(ar_path):
        with open(ar_path) as f:
            result["ar_aging"] = json.load(f)
    else:
        result["ar_aging"] = []

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
```

- [ ] **Step 2: Verify save/load round-trip**

```powershell
python -c "
from app.models import ARAgingRow, QualityCheck
from app.storage.runs import save_run, load_run, new_run_id
import os, shutil

run_id = new_run_id()
ar = [ARAgingRow('Prop A','PM','f.xlsx','Tenant Rent',2024,3,10000,8000,500,300,200,100,-50)]
save_run(run_id, {'portfolio_name':'Test'}, [], [], [], [], ar_rows=ar)
data = load_run(run_id)
assert len(data['ar_aging']) == 1
assert data['ar_aging'][0]['receivable_type'] == 'Tenant Rent'
shutil.rmtree(os.path.join('runs', run_id))
print('OK')
"
```

Expected: `OK`

- [ ] **Step 3: Verify backward compat (run without ar_aging.json returns empty list)**

```powershell
python -c "
from app.storage.runs import save_run, load_run, new_run_id
import os, shutil

run_id = new_run_id()
save_run(run_id, {'portfolio_name':'Test'}, [], [], [], [])
data = load_run(run_id)
assert data['ar_aging'] == [], f'Expected [], got {data[\"ar_aging\"]}'
shutil.rmtree(os.path.join('runs', run_id))
print('OK - backward compat confirmed')
"
```

Expected: `OK - backward compat confirmed`

- [ ] **Step 4: Commit**

```powershell
git add app/storage/runs.py
git commit -m "feat: save/load ar_aging.json in run storage (backward compat)"
```

---

## Task 4: Upload Pipeline + Drop Zone UI

**Files:**
- Modify: `app/templates/upload.html`
- Modify: `app/routes/upload.py`

- [ ] **Step 1: Add AR Aging drop zone to `app/templates/upload.html`**

Insert the following new `<section>` block immediately **after** the closing `</section>` of the Physical Occupancy card and **before** the opening `<section>` of the Analysis Settings card (between lines 50 and 52 of the current file):

```html
  <section class="card">
    <h2>AR Aging Reports (Optional)</h2>
    <div class="field">
      <label>AR Aging Workbooks (.xlsx)</label>
      <div class="drop-zone" id="ar-drop-zone">
        <input type="file" id="ar_aging_files" name="ar_aging_files"
               class="drop-zone-input" multiple accept=".xlsx,.xls">
        <div class="drop-zone-body">
          <div class="drop-zone-icon">📋</div>
          <div class="drop-zone-text">Drag &amp; drop AR aging files here</div>
          <div class="drop-zone-sub">or <span class="browse-link" data-for="ar_aging_files">browse files</span></div>
        </div>
        <ul class="file-list" id="ar-file-list"></ul>
      </div>
      <span class="hint">Files named <code>PMC_AR Aging_Tenant Rent_MM_YYYY.xlsx</code> or <code>PMC_AR Aging_Subsidy_MM_YYYY.xlsx</code>. Multiple files accepted — one per period per type per PM company.</span>
    </div>
  </section>
```

Also add the drop zone wiring at the bottom of the `<script>` block, just before the closing `})();` line:

```javascript
  initDropZone('ar-drop-zone', 'ar_aging_files', 'ar-file-list');
```

- [ ] **Step 2: Update `app/routes/upload.py`**

Replace the entire file:

```python
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
    eco_occ_target = float(request.form.get("eco_occ_target", ECO_OCC_TARGET * 100)) / 100
    pm_names_raw   = request.form.get("pm_names", "").strip()
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

    # Apply PROPERTY_NAME_MAP to AR rows (parser already does it, but apply again for safety)
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
            k.is_below_eco_occ_target = k.eco_occ_pct < eco_occ_target

    # ── Build workbooks ───────────────────────────────────────────────────────
    run_id  = new_run_id()
    run_dir = os.path.join("runs", run_id)
    os.makedirs(run_dir, exist_ok=True)

    safe_name   = "".join(c for c in portfolio_name if c.isalnum() or c in " _-").strip()
    main_path   = os.path.join(run_dir, f"{safe_name} Property Analysis.xlsx")
    backup_path = os.path.join(run_dir, f"{safe_name} Property Analysis backup.xlsx")

    build_main_workbook(kpis, portfolio_name, main_path, eco_occ_target,
                        ar_rows=ar_rows if ar_rows else None)
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
```

- [ ] **Step 3: Verify the app still starts cleanly**

```powershell
python -c "from app import create_app; app = create_app(); print('OK')"
```

Expected: `OK`

- [ ] **Step 4: Commit**

```powershell
git add app/templates/upload.html app/routes/upload.py
git commit -m "feat: AR aging upload drop zone and pipeline integration"
```

---

## Task 5: Results Route AR Helpers

**Files:**
- Modify: `app/routes/results.py`

- [ ] **Step 1: Replace `app/routes/results.py` with the updated version**

```python
from flask import Blueprint, render_template, abort
from app.storage.runs import load_run
from config import ECO_OCC_TARGET

bp = Blueprint("results", __name__)

# Tooltip text for KPI labels in the web dashboard
_KPI_TOOLTIPS = {
    "Actual Income":      "Total Income = GPR + Other Income − Vacancy − Concessions − Bad Debt (Effective Gross Income)",
    "Budget Income":      "Budgeted Total Income for the period",
    "Income Variance":    "Actual Income − Budget Income. Positive = favorable (above budget)",
    "Income Variance %":  "Income Variance / Budget Income",
    "Actual Expenses":    "Sum of all Operating Expense accounts. Excludes depreciation, debt service, reserves",
    "Budget Expenses":    "Budgeted Operating Expenses for the period",
    "Expense Variance":   "Actual Expenses − Budget Expenses. Negative = favorable (under budget)",
    "Expense Variance %": "Expense Variance / Budget Expenses",
    "Actual NOI":         "NOI = Total Income − Total Operating Expenses",
    "Budget NOI":         "Budget NOI = Budget Income − Budget Expenses",
    "NOI Variance":       "NOI Variance = Actual NOI − Budget NOI. Positive = favorable",
    "NOI Variance %":     "NOI Variance / |Budget NOI|. Absolute denominator handles sign flips when budget NOI is negative",
    "GPR":                "Gross Potential Rent — total scheduled rent before any deductions",
    "Vacancy":            "Vacancy loss — rent foregone from unoccupied units",
    "Concessions":        "Move-in specials and rent concessions",
    "Bad Debt":           "Collection losses and write-offs",
    "Net Collectible":    "GPR − Vacancy − Concessions − Bad Debt",
    "Eco Occ %":          "Economic Occupancy % = Net Collectible / GPR",
    "Budget Eco Occ %":   "Budget Economic Occupancy % = Budget Net Collectible / Budget GPR",
    "Eco Occ Variance":   "Actual Eco Occ % − Budget Eco Occ %",
    "Physical Occ %":     "Physical Occ % = Occupied Units / Total Units. Sourced from Physical Occupancy Report",
    "Leakage Gap":        "Physical Occ % − Economic Occ %. Positive = units occupied but rent not being fully collected",
    "Income/Unit":        "Actual Income / Total Units (from Physical Occupancy Report)",
    "Expense/Unit":       "Actual Expenses / Total Units (from Physical Occupancy Report)",
    "NOI/Unit":           "Actual NOI / Total Units (from Physical Occupancy Report)",
}

_SUMMARY_KPI_DEFINITIONS = [
    ("Actual Income",      "actual_income",        "currency", None),
    ("Budget Income",      "budget_income",        "currency", None),
    ("Income Variance",    "income_variance",      "currency", True),
    ("Income Variance %",  "income_variance_pct",  "pct",      True),
    None,
    ("Actual Expenses",    "actual_expenses",      "currency", None),
    ("Budget Expenses",    "budget_expenses",      "currency", None),
    ("Expense Variance",   "expense_variance",     "currency", False),
    ("Expense Variance %", "expense_variance_pct", "pct",      False),
    None,
    ("Actual NOI",         "actual_noi",           "currency", None),
    ("Budget NOI",         "budget_noi",           "currency", None),
    ("NOI Variance",       "noi_variance",         "currency", True),
    ("NOI Variance %",     "noi_variance_pct",     "pct",      True),
    None,
    ("GPR",                "gpr",                  "currency", None),
    ("Vacancy",            "vacancy",              "currency", None),
    ("Concessions",        "concessions",          "currency", None),
    ("Bad Debt",           "bad_debt",             "currency", None),
    ("Net Collectible",    "net_collectible",      "currency", None),
    ("Eco Occ %",          "eco_occ_pct",          "pct",      None),
    ("Budget Eco Occ %",   "budget_eco_occ_pct",   "pct",      None),
    ("Eco Occ Variance",   "eco_occ_variance",     "pct",      True),
    None,
    ("Physical Occ %",     "physical_occ_pct",     "pct",      None),
    ("Leakage Gap",        "leakage_gap",          "pct",      False),
    None,
    ("Income/Unit",        "income_per_unit",      "currency", None),
    ("Expense/Unit",       "expense_per_unit",     "currency", None),
    ("NOI/Unit",           "noi_per_unit",         "currency", None),
]


# ── Quarter helpers ────────────────────────────────────────────────────────────

def _month_to_quarter(month: int) -> int:
    return (month - 1) // 3 + 1


def _quarter_label(year: int, quarter: int) -> str:
    return f"Q{quarter} - {year}"


# ── AR Aging helpers ───────────────────────────────────────────────────────────

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


def _ar_period_label(year: int, month: int) -> str:
    """Return display label e.g. 'Mar-2024'."""
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


def _agg_ar(ar_rows: list[dict], receivable_type: str, year: int, month: int) -> dict | None:
    """Aggregate AR rows for a specific type/period. Returns None if no matching rows."""
    rows = [r for r in ar_rows
            if r["receivable_type"] == receivable_type
            and r["year"] == year
            and r["month"] == month]
    if not rows:
        return None
    charge_amount = sum(r["charge_amount"] for r in rows)
    total_overdue = sum(r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed":   sum(r["current_owed"] for r in rows),
        "prepayments":    sum(r["prepayments"] for r in rows),
        "pct_overdue":    total_overdue / charge_amount if charge_amount > 0 else None,
        "property_count": len({r["property_name"] for r in rows}),
    }


def _agg_ar_for_prop(ar_rows: list[dict], property_name: str,
                     receivable_type: str, year: int, month: int) -> dict | None:
    """Aggregate AR rows for a specific property/type/period."""
    rows = [r for r in ar_rows
            if r["property_name"] == property_name
            and r["receivable_type"] == receivable_type
            and r["year"] == year
            and r["month"] == month]
    if not rows:
        return None
    charge = sum(r["charge_amount"] for r in rows)
    overdue = sum(r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"] for r in rows)
    return {
        "current_owed": sum(r["current_owed"] for r in rows),
        "prepayments":  sum(r["prepayments"] for r in rows),
        "pct_overdue":  overdue / charge if charge > 0 else None,
    }


def _ar_yoy_delta(curr: dict, prev: dict) -> dict:
    """Compute YoY delta between two period aggregates."""
    pct_delta = None
    if curr.get("pct_overdue") is not None and prev.get("pct_overdue") is not None:
        pct_delta = curr["pct_overdue"] - prev["pct_overdue"]
    return {
        "current_owed_delta": curr["current_owed"] - prev["current_owed"],
        "prepayments_delta":  curr["prepayments"] - prev["prepayments"],
        "pct_overdue_delta":  pct_delta,
    }


def _pct_delta(curr: dict | None, prev: dict | None) -> float | None:
    """Return pct_overdue delta between two property aggregates, or None."""
    if curr and prev:
        c = curr.get("pct_overdue")
        p = prev.get("pct_overdue")
        if c is not None and p is not None:
            return c - p
    return None


# ── KPI aggregation ────────────────────────────────────────────────────────────

def _agg_kpis(kpi_dicts: list[dict]) -> dict:
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

    actual_noi = (actual_income - actual_expenses) if (actual_income is not None and actual_expenses is not None) else None
    budget_noi = (budget_income - budget_expenses) if (budget_income is not None and budget_expenses is not None) else None

    net_coll = (gpr - (vacancy or 0) - (concessions or 0) - (bad_debt or 0)) if gpr is not None else None
    eco_occ  = (net_coll / gpr) if (net_coll is not None and gpr) else None

    bud_eco_vals = [k["budget_eco_occ_pct"] for k in kpi_dicts if k.get("budget_eco_occ_pct") is not None]
    bud_eco      = sum(bud_eco_vals) / len(bud_eco_vals) if bud_eco_vals else None
    eco_occ_var  = (eco_occ - bud_eco) if (eco_occ is not None and bud_eco is not None) else None

    noi_var     = (actual_noi - budget_noi) if (actual_noi is not None and budget_noi is not None) else None
    noi_var_pct = (noi_var / abs(budget_noi)) if (noi_var is not None and budget_noi) else None

    _paired = [
        (k["occupied_units"], k["total_units"])
        for k in kpi_dicts
        if k.get("occupied_units") is not None and k.get("total_units") is not None
    ]
    if _paired:
        _occ_sum   = sum(p[0] for p in _paired)
        _total_sum = sum(p[1] for p in _paired)
        phys_occ   = _occ_sum / _total_sum if _total_sum > 0 else None
    else:
        phys_occ = None

    total_units = next((k["total_units"] for k in kpi_dicts if k.get("total_units") is not None), None)

    income_pu  = (actual_income   / total_units) if (actual_income   is not None and total_units) else None
    expense_pu = (actual_expenses / total_units) if (actual_expenses is not None and total_units) else None
    noi_pu     = (actual_noi      / total_units) if (actual_noi      is not None and total_units) else None

    def _safe_pct(num, denom):
        if num is None or denom is None or denom == 0:
            return None
        return num / abs(denom)

    return dict(
        actual_income=actual_income,
        budget_income=budget_income,
        income_variance=(actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None,
        income_variance_pct=_safe_pct(
            (actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None,
            budget_income,
        ),
        actual_expenses=actual_expenses,
        budget_expenses=budget_expenses,
        expense_variance=(actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None,
        expense_variance_pct=_safe_pct(
            (actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None,
            budget_expenses,
        ),
        actual_noi=actual_noi,
        budget_noi=budget_noi,
        noi_variance=noi_var,
        noi_variance_pct=noi_var_pct,
        eco_occ_pct=eco_occ,
        budget_eco_occ_pct=bud_eco,
        eco_occ_variance=eco_occ_var,
        physical_occ_pct=phys_occ,
        leakage_gap=(phys_occ - eco_occ) if (phys_occ is not None and eco_occ is not None) else None,
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


# ── Route ──────────────────────────────────────────────────────────────────────

@bp.route("/results/<run_id>")
def show(run_id):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    meta   = data["metadata"]
    kpis   = data["kpis"]
    checks = data["quality_checks"]

    portfolio_name = meta.get("portfolio_name", "Portfolio")
    eco_occ_target = meta.get("eco_occ_target", ECO_OCC_TARGET)
    years  = meta.get("years", [])
    props  = meta.get("properties", [])

    # ── Quarter-period aggregation ─────────────────────────────────────────────
    all_quarters: set[tuple] = set()
    for k in kpis:
        if not k.get("is_carveout") and k.get("year") and k.get("month"):
            all_quarters.add((k["year"], _month_to_quarter(k["month"])))
    sorted_quarters = sorted(all_quarters)
    period_labels = [_quarter_label(yr, q) for (yr, q) in sorted_quarters]

    period_aggs: dict[str, dict] = {}
    period_property_counts: dict[str, int] = {}
    for (yr, q) in sorted_quarters:
        months = {(q - 1) * 3 + 1, (q - 1) * 3 + 2, (q - 1) * 3 + 3}
        q_kpis = [
            k for k in kpis
            if k.get("year") == yr and k.get("month") in months and not k.get("is_carveout")
        ]
        lbl = _quarter_label(yr, q)
        period_aggs[lbl] = _agg_kpis(q_kpis)
        period_property_counts[lbl] = len({k["property_name"] for k in q_kpis})

    latest_period_label = period_labels[-1] if period_labels else ""

    summary_kpi_rows = []
    for defn in _SUMMARY_KPI_DEFINITIONS:
        if defn is None:
            summary_kpi_rows.append({"sep": True})
            continue
        label, key, fmt, fav = defn
        values = [period_aggs.get(lbl, {}).get(key) for lbl in period_labels]
        summary_kpi_rows.append({
            "sep": False,
            "label": label,
            "key": key,
            "fmt": fmt,
            "favorable_positive": fav,
            "tooltip": _KPI_TOOLTIPS.get(label, ""),
            "period_values": values,
        })

    # ── Property table ─────────────────────────────────────────────────────────
    latest_yr = max(years) if years else None
    prop_rows = []
    for prop in sorted(props):
        prop_kpis = [
            k for k in kpis
            if k.get("property_name") == prop and k.get("year") == latest_yr
        ]
        if prop_kpis:
            agg = _agg_kpis(prop_kpis)
            agg["property_name"] = prop
            agg["pm_name"] = prop_kpis[0].get("pm_name", "")
            agg["top_noi_driver_1"] = prop_kpis[0].get("top_noi_driver_1", "")
            agg["top_noi_driver_2"] = prop_kpis[0].get("top_noi_driver_2", "")
            prop_rows.append(agg)

    # ── AR Aging section ───────────────────────────────────────────────────────
    ar_rows = data.get("ar_aging", [])
    ar_summary: dict = {}
    ar_prop_rows: list = []
    ar_latest_period_label: str = ""

    if ar_rows:
        for rtype in ["Tenant Rent", "Subsidy"]:
            periods = sorted({(r["year"], r["month"]) for r in ar_rows
                              if r["receivable_type"] == rtype})
            if not periods:
                continue
            period_set = set(periods)
            cols = []
            for (yr, mo) in periods:
                agg = _agg_ar(ar_rows, rtype, yr, mo)
                cols.append({
                    "type":         "period",
                    "label":        _ar_period_label(yr, mo),
                    "year": yr, "month": mo,
                    "current_owed": agg["current_owed"] if agg else None,
                    "prepayments":  agg["prepayments"]  if agg else None,
                    "pct_overdue":  agg["pct_overdue"]  if agg else None,
                })
                if (yr - 1, mo) in period_set:
                    curr = agg
                    prev = _agg_ar(ar_rows, rtype, yr - 1, mo)
                    if curr and prev:
                        delta = _ar_yoy_delta(curr, prev)
                    else:
                        delta = {"current_owed_delta": None,
                                 "prepayments_delta": None,
                                 "pct_overdue_delta": None}
                    cols.append({
                        "type":         "yoy",
                        "label":        "YoY Δ",
                        "year": yr, "month": mo,
                        "current_owed": delta["current_owed_delta"],
                        "prepayments":  delta["prepayments_delta"],
                        "pct_overdue":  delta["pct_overdue_delta"],
                    })

            prop_count = len({r["property_name"] for r in ar_rows
                              if r["receivable_type"] == rtype})
            ar_summary[rtype] = {"property_count": prop_count, "cols": cols}

        # Property-level AR table — latest period across all types
        latest_yr_ar, latest_mo_ar = max((r["year"], r["month"]) for r in ar_rows)
        ar_latest_period_label = _ar_period_label(latest_yr_ar, latest_mo_ar)

        all_latest_props = sorted({r["property_name"] for r in ar_rows
                                   if r["year"] == latest_yr_ar and r["month"] == latest_mo_ar})

        for prop in all_latest_props:
            pm = next((r["pm_name"] for r in ar_rows
                       if r["property_name"] == prop
                       and r["year"] == latest_yr_ar
                       and r["month"] == latest_mo_ar), "")

            tr_curr  = _agg_ar_for_prop(ar_rows, prop, "Tenant Rent", latest_yr_ar, latest_mo_ar)
            tr_prev  = _agg_ar_for_prop(ar_rows, prop, "Tenant Rent", latest_yr_ar - 1, latest_mo_ar)
            sub_curr = _agg_ar_for_prop(ar_rows, prop, "Subsidy",     latest_yr_ar, latest_mo_ar)
            sub_prev = _agg_ar_for_prop(ar_rows, prop, "Subsidy",     latest_yr_ar - 1, latest_mo_ar)

            ar_prop_rows.append({
                "property_name":   prop,
                "pm_name":         pm,
                "tr_current_owed":  tr_curr["current_owed"]  if tr_curr else None,
                "tr_prepayments":   tr_curr["prepayments"]   if tr_curr else None,
                "tr_pct_overdue":   tr_curr["pct_overdue"]   if tr_curr else None,
                "tr_yoy_pct_delta": _pct_delta(tr_curr, tr_prev),
                "sub_current_owed":  sub_curr["current_owed"]  if sub_curr else None,
                "sub_prepayments":   sub_curr["prepayments"]   if sub_curr else None,
                "sub_pct_overdue":   sub_curr["pct_overdue"]   if sub_curr else None,
                "sub_yoy_pct_delta": _pct_delta(sub_curr, sub_prev),
            })

        # Sort by Tenant Rent current_owed descending
        ar_prop_rows.sort(key=lambda r: (r["tr_current_owed"] or 0), reverse=True)

    return render_template(
        "dashboard.html",
        run_id=run_id,
        meta=meta,
        kpis=kpis,
        period_labels=period_labels,
        period_property_counts=period_property_counts,
        summary_kpi_rows=summary_kpi_rows,
        prop_rows=prop_rows,
        quality_checks=checks,
        portfolio_name=portfolio_name,
        eco_occ_target=eco_occ_target,
        latest_period_label=latest_period_label,
        years=years,
        properties=props,
        num_properties=len(props),
        # AR Aging context
        ar_summary=ar_summary,
        ar_prop_rows=ar_prop_rows,
        ar_latest_period_label=ar_latest_period_label,
    )
```

- [ ] **Step 2: Verify import and route registration**

```powershell
python -c "from app.routes.results import show, _agg_ar, _ar_yoy_delta, _ar_period_label; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```powershell
git add app/routes/results.py
git commit -m "feat: results route AR Aging helpers and context variables"
```

---

## Task 6: Dashboard Template AR Section

**Files:**
- Modify: `app/templates/dashboard.html`

- [ ] **Step 1: Add AR Aging section to `app/templates/dashboard.html`**

Insert the following block immediately **before** the `<!-- Quality Checks -->` comment (before the `{% if quality_checks %}` line):

```html
<!-- AR Aging Section -->
{% if ar_summary %}
{% for rtype in ["Tenant Rent", "Subsidy"] %}
{% if ar_summary[rtype] is defined %}
{% set summary = ar_summary[rtype] %}
<section class="card">
  <h2>Portfolio AR Summary — {{ rtype }} ({{ summary.property_count }} Properties)</h2>
  <div class="table-scroll">
  <table class="kpi-table kpi-transposed">
    <thead>
      <tr>
        <th class="kpi-label-col">Metric</th>
        {% for col in summary.cols %}
        <th class="{{ 'yoy-col' if col.type == 'yoy' else '' }}">{{ col.label }}</th>
        {% endfor %}
      </tr>
    </thead>
    <tbody>
      <!-- Current Owed -->
      <tr>
        <td class="kpi-label" title="Current period AR balance outstanding">Current Owed</td>
        {% for col in summary.cols %}
        {% if col.type == 'period' %}
        <td class="currency">{{ col.current_owed | currency }}</td>
        {% else %}
        <td class="currency {{ 'neg' if (col.current_owed or 0) > 0 else ('pos' if (col.current_owed or 0) < 0 else '') }}">
          {{ col.current_owed | currency }}</td>
        {% endif %}
        {% endfor %}
      </tr>
      <!-- Pre-payments -->
      <tr>
        <td class="kpi-label" title="Credit balances (negative values = credits held)">Pre-payments</td>
        {% for col in summary.cols %}
        <td class="currency">{{ col.prepayments | currency }}</td>
        {% endfor %}
      </tr>
      <!-- % >30 Days -->
      <tr>
        <td class="kpi-label" title="(31-60 + 61-90 + Over 90) / Charge Amount">% &gt;30 Days</td>
        {% for col in summary.cols %}
        {% if col.type == 'period' %}
        <td class="pct">{{ col.pct_overdue | pct }}</td>
        {% else %}
        <td class="pct {{ 'neg' if (col.pct_overdue or 0) > 0 else ('pos' if (col.pct_overdue or 0) < 0 else '') }}">
          {{ col.pct_overdue | pct }}</td>
        {% endif %}
        {% endfor %}
      </tr>
    </tbody>
  </table>
  </div>
</section>
{% endif %}
{% endfor %}

{% if ar_prop_rows %}
<section class="card">
  <h2>Property AR Analysis — {{ ar_latest_period_label }}</h2>
  <div class="table-scroll">
  <table class="data-table">
    <thead>
      <tr>
        <th>Property</th>
        <th>PM</th>
        <th title="Tenant Rent Current Owed">TR Current Owed</th>
        <th title="Tenant Rent Pre-payments (credits)">TR Pre-payments</th>
        <th title="Tenant Rent % of charge amount >30 days overdue">TR % &gt;30</th>
        <th title="YoY change in Tenant Rent % >30 Days (pp)">TR YoY Δ % &gt;30</th>
        <th title="Subsidy Current Owed">Sub Current Owed</th>
        <th title="Subsidy Pre-payments (credits)">Sub Pre-payments</th>
        <th title="Subsidy % of charge amount >30 days overdue">Sub % &gt;30</th>
        <th title="YoY change in Subsidy % >30 Days (pp)">Sub YoY Δ % &gt;30</th>
      </tr>
    </thead>
    <tbody>
      {% for r in ar_prop_rows %}
      <tr>
        <td>{{ r.property_name }}</td>
        <td>{{ r.pm_name }}</td>
        <td class="currency">{{ r.tr_current_owed | currency }}</td>
        <td class="currency">{{ r.tr_prepayments | currency }}</td>
        <td class="pct">{{ r.tr_pct_overdue | pct }}</td>
        <td class="pct {{ 'neg' if (r.tr_yoy_pct_delta or 0) > 0 else ('pos' if (r.tr_yoy_pct_delta or 0) < 0 else '') }}">
          {{ r.tr_yoy_pct_delta | pct }}</td>
        <td class="currency">{{ r.sub_current_owed | currency }}</td>
        <td class="currency">{{ r.sub_prepayments | currency }}</td>
        <td class="pct">{{ r.sub_pct_overdue | pct }}</td>
        <td class="pct {{ 'neg' if (r.sub_yoy_pct_delta or 0) > 0 else ('pos' if (r.sub_yoy_pct_delta or 0) < 0 else '') }}">
          {{ r.sub_yoy_pct_delta | pct }}</td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
  </div>
</section>
{% endif %}
{% endif %}
```

- [ ] **Step 2: Commit**

```powershell
git add app/templates/dashboard.html
git commit -m "feat: dashboard AR Aging portfolio summary and property table sections"
```

---

## Task 7: Property Detail Page AR Section

**Files:**
- Modify: `app/routes/property_detail.py`
- Modify: `app/templates/property_detail.html`

- [ ] **Step 1: Update `app/routes/property_detail.py`**

Replace the entire file:

```python
from flask import Blueprint, render_template, abort
from app.storage.runs import load_run

bp = Blueprint("property_detail", __name__)

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


@bp.route("/results/<run_id>/property/<property_name>")
def show(run_id, property_name):
    try:
        data = load_run(run_id)
    except FileNotFoundError:
        abort(404)

    prop_kpis = [k for k in data["kpis"] if k["property_name"] == property_name]
    if not prop_kpis:
        abort(404)

    # Build AR Aging rows for this property, sorted by (receivable_type, year, month)
    ar_all = data.get("ar_aging", [])
    prop_ar_rows = sorted(
        [r for r in ar_all if r["property_name"] == property_name],
        key=lambda r: (r["receivable_type"], r["year"], r["month"])
    )
    # Augment each row with computed fields (not stored in JSON — @property on dataclass)
    for r in prop_ar_rows:
        total_overdue = r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"]
        r["pct_overdue"]   = total_overdue / r["charge_amount"] if r["charge_amount"] > 0 else None
        r["total_overdue"] = total_overdue
        r["period_label"]  = f"{_MONTH_ABBR.get(r['month'], str(r['month']))}-{r['year']}"

    return render_template(
        "property_detail.html",
        run_id=run_id,
        property_name=property_name,
        kpis=prop_kpis,
        meta=data["metadata"],
        prop_ar_rows=prop_ar_rows,
    )
```

- [ ] **Step 2: Add AR Aging section to `app/templates/property_detail.html`**

Append the following immediately **after** the closing `</section>` of the Monthly KPIs card (after line 58 of the current file):

```html
{% if prop_ar_rows %}
<section class="card">
  <h2>AR Aging</h2>
  {% for rtype in ["Tenant Rent", "Subsidy"] %}
  {% set type_rows = prop_ar_rows | selectattr("receivable_type", "equalto", rtype) | list %}
  {% if type_rows %}
  <h3 style="margin:12px 0 8px; font-size:14px; color:#2e75b6;">{{ rtype }}</h3>
  <div class="table-scroll">
  <table class="data-table">
    <thead>
      <tr>
        <th>Period</th>
        <th title="Total charges outstanding">Charge Amt</th>
        <th title="Current period balance">Current Owed</th>
        <th title="0–30 day aging bucket">0–30 Owed</th>
        <th title="31–60 day aging bucket">31–60 Owed</th>
        <th title="61–90 day aging bucket">61–90 Owed</th>
        <th title="90+ day aging bucket">Over 90 Owed</th>
        <th title="Credit balances (negative = credits)">Pre-payments</th>
        <th title="(31-60 + 61-90 + Over 90) / Charge Amount">% &gt;30 Days</th>
      </tr>
    </thead>
    <tbody>
      {% for r in type_rows %}
      <tr>
        <td>{{ r.period_label }}</td>
        <td class="currency">{{ r.charge_amount | currency }}</td>
        <td class="currency">{{ r.current_owed | currency }}</td>
        <td class="currency">{{ r.owed_0_30 | currency }}</td>
        <td class="currency">{{ r.owed_31_60 | currency }}</td>
        <td class="currency">{{ r.owed_61_90 | currency }}</td>
        <td class="currency">{{ r.owed_over_90 | currency }}</td>
        <td class="currency">{{ r.prepayments | currency }}</td>
        <td class="pct">{{ r.pct_overdue | pct }}</td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
  </div>
  {% endif %}
  {% endfor %}
</section>
{% endif %}
```

- [ ] **Step 3: Commit**

```powershell
git add app/routes/property_detail.py app/templates/property_detail.html
git commit -m "feat: property detail AR Aging section (all periods, full bucket breakdown)"
```

---

## Task 8: Main Workbook AR Aging Tab

**Files:**
- Modify: `app/exporter/main_workbook.py`

- [ ] **Step 1: Add `_ar_period_label` helper and update `build_main_workbook` signature**

In `app/exporter/main_workbook.py`, make the following changes:

**a.** Add `_ar_period_label` helper function after the `_get_sorted_quarters` function (after line 34):

```python
def _ar_period_label(year: int, month: int) -> str:
    """Return AR period display label, e.g. 'Mar-2024'."""
    _ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
             7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    return f"{_ABBR.get(month, str(month))}-{year}"
```

**b.** Update `build_main_workbook` signature and body. Replace:

```python
def build_main_workbook(
    kpis: list[PropertyPeriodKPIs],
    portfolio_name: str,
    output_path: str,
    eco_occ_target: float = ECO_OCC_TARGET,
) -> str:
    """Builds main workbook at output_path. Returns path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_dashboard(wb, kpis, portfolio_name, eco_occ_target)
    _build_property_analysis(wb, kpis, portfolio_name, eco_occ_target)
    _build_monthly_kpis(wb, kpis)

    wb.save(output_path)
    return output_path
```

With:

```python
def build_main_workbook(
    kpis: list[PropertyPeriodKPIs],
    portfolio_name: str,
    output_path: str,
    eco_occ_target: float = ECO_OCC_TARGET,
    ar_rows: list | None = None,
) -> str:
    """Builds main workbook at output_path. Returns path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_dashboard(wb, kpis, portfolio_name, eco_occ_target)
    _build_property_analysis(wb, kpis, portfolio_name, eco_occ_target)
    _build_monthly_kpis(wb, kpis)
    _build_ar_aging(wb, ar_rows, portfolio_name)

    wb.save(output_path)
    return output_path
```

**c.** Append the `_build_ar_aging` function at the end of `app/exporter/main_workbook.py` (after `_autofit_columns`):

```python
# ─── AR Aging tab ─────────────────────────────────────────────────────────────

def _build_ar_aging(wb, ar_rows, portfolio_name):
    """Build the 'AR Aging' tab with 4 blocks: portfolio summary × 2 types + property analysis × 2."""
    ws = wb.create_sheet("AR Aging")

    if not ar_rows:
        ws.cell(1, 1, "No AR Aging data uploaded — upload AR Aging files on the upload page.")
        ws.cell(1, 1).font = BOLD_FONT
        return

    def _agg(rtype, yr, mo):
        """Aggregate ar_rows for a specific receivable type / period."""
        rows = [r for r in ar_rows
                if r.receivable_type == rtype and r.year == yr and r.month == mo]
        if not rows:
            return None
        charge  = sum(r.charge_amount for r in rows)
        overdue = sum(r.total_overdue for r in rows)
        return {
            "current_owed":   sum(r.current_owed for r in rows),
            "prepayments":    sum(r.prepayments  for r in rows),
            "pct_overdue":    overdue / charge if charge > 0 else None,
            "property_count": len({r.property_name for r in rows}),
        }

    row = 1

    # ── Blocks 1 & 2: Portfolio AR Summary ──────────────────────────────────
    for rtype in ["Tenant Rent", "Subsidy"]:
        periods = sorted({(r.year, r.month) for r in ar_rows if r.receivable_type == rtype})
        if not periods:
            row += 2
            continue

        period_set = set(periods)
        prop_count = len({r.property_name for r in ar_rows if r.receivable_type == rtype})

        # Build col_specs: ("period"|"yoy", year, month)
        col_specs = []
        for (yr, mo) in periods:
            col_specs.append(("period", yr, mo))
            if (yr - 1, mo) in period_set:
                col_specs.append(("yoy", yr, mo))

        # Section header
        ws.cell(row, 1, f"Portfolio AR Summary — {rtype} ({prop_count} Properties)")
        ws.cell(row, 1).font = BOLD_FONT
        row += 1

        # Column headers
        ws.cell(row, 1, "Metric")
        for ci, (typ, yr, mo) in enumerate(col_specs, 2):
            if typ == "period":
                ws.cell(row, ci, _ar_period_label(yr, mo))
            else:
                ws.cell(row, ci, f"YoY Δ vs {_ar_period_label(yr - 1, mo)}")
        style_header_row(ws, row, 1 + len(col_specs))
        row += 1

        # Metric rows: Current Owed, Pre-payments, % >30 Days
        for metric_label, metric_key, fmt, unfav_is_increase in [
            ("Current Owed", "current_owed", CURRENCY_FMT, True),
            ("Pre-payments", "prepayments",  CURRENCY_FMT, None),  # no color for prepayments
            ("% >30 Days",   "pct_overdue",  PCT_FMT,      True),
        ]:
            ws.cell(row, 1, metric_label)
            for ci, (typ, yr, mo) in enumerate(col_specs, 2):
                if typ == "period":
                    agg = _agg(rtype, yr, mo)
                    val = agg[metric_key] if agg else None
                    if val is not None:
                        _c(ws, row, ci, val, fmt)
                    else:
                        ws.cell(row, ci, None)
                else:  # yoy delta
                    curr_agg = _agg(rtype, yr, mo)
                    prev_agg = _agg(rtype, yr - 1, mo)
                    if curr_agg and prev_agg:
                        cv = curr_agg.get(metric_key)
                        pv = prev_agg.get(metric_key)
                        delta = (cv - pv) if (cv is not None and pv is not None) else None
                    else:
                        delta = None
                    cell = ws.cell(row, ci, delta)
                    if delta is not None:
                        cell.number_format = fmt
                        if unfav_is_increase is not None:
                            # unfav_is_increase=True: increase is bad (red), decrease is good (green)
                            apply_variance_fill(cell, delta, favorable_is_positive=not unfav_is_increase)
            row += 1

        row += 3  # blank separator before next block

    # ── Blocks 3 & 4: Property AR Analysis ──────────────────────────────────
    latest_yr, latest_mo = max((r.year, r.month) for r in ar_rows)
    latest_lbl = _ar_period_label(latest_yr, latest_mo)

    for rtype in ["Tenant Rent", "Subsidy"]:
        latest_rows = [r for r in ar_rows
                       if r.receivable_type == rtype
                       and r.year == latest_yr and r.month == latest_mo]
        if not latest_rows:
            continue

        ws.cell(row, 1, f"Property AR Analysis — {rtype} (As of {latest_lbl})")
        ws.cell(row, 1).font = BOLD_FONT
        row += 1

        prop_headers = ["Property", "PM", "Current Owed", "Pre-payments",
                        "% >30 Days", "YoY $ Δ (Current Owed)", "YoY Δ (% >30 Days)"]
        for ci, h in enumerate(prop_headers, 1):
            ws.cell(row, ci, h)
        style_header_row(ws, row, len(prop_headers))
        prop_header_row = row
        row += 1

        # Group latest-period rows by property
        by_prop: dict[str, list] = {}
        for r in latest_rows:
            by_prop.setdefault(r.property_name, []).append(r)

        for prop_name in sorted(by_prop):
            group    = by_prop[prop_name]
            pm       = group[0].pm_name
            charge   = sum(r.charge_amount for r in group)
            overdue  = sum(r.total_overdue for r in group)
            c_owed   = sum(r.current_owed  for r in group)
            prepy    = sum(r.prepayments   for r in group)
            pct_ov   = overdue / charge if charge > 0 else None

            # Prior-year same-month for this property/type
            prev_grp = [r for r in ar_rows
                        if r.receivable_type == rtype
                        and r.property_name == prop_name
                        and r.year == latest_yr - 1 and r.month == latest_mo]
            yoy_c = None
            yoy_p = None
            if prev_grp:
                prev_charge  = sum(r.charge_amount for r in prev_grp)
                prev_overdue = sum(r.total_overdue  for r in prev_grp)
                prev_c_owed  = sum(r.current_owed   for r in prev_grp)
                prev_pct     = prev_overdue / prev_charge if prev_charge > 0 else None
                yoy_c = c_owed - prev_c_owed
                if pct_ov is not None and prev_pct is not None:
                    yoy_p = pct_ov - prev_pct

            ws.cell(row, 1, prop_name)
            ws.cell(row, 2, pm)
            _c(ws, row, 3, c_owed, CURRENCY_FMT)
            _c(ws, row, 4, prepy,  CURRENCY_FMT)
            if pct_ov is not None:
                _c(ws, row, 5, pct_ov, PCT_FMT)

            co_cell = ws.cell(row, 6, yoy_c)
            if yoy_c is not None:
                co_cell.number_format = CURRENCY_FMT
                apply_variance_fill(co_cell, yoy_c, favorable_is_positive=False)

            pct_cell = ws.cell(row, 7, yoy_p)
            if yoy_p is not None:
                pct_cell.number_format = PCT_FMT
                apply_variance_fill(pct_cell, yoy_p, favorable_is_positive=False)

            row += 1

        ws.auto_filter.ref = (
            f"A{prop_header_row}:{get_column_letter(len(prop_headers))}{row - 1}"
        )
        row += 3

    _autofit_columns(ws)
```

- [ ] **Step 2: Smoke test the main workbook builder**

```powershell
python -c "
import os, tempfile
from app.models import PropertyPeriodKPIs, ARAgingRow
from app.exporter.main_workbook import build_main_workbook

k = PropertyPeriodKPIs('Test Prop','PM Co',2024,1,'Jan')
k.actual_income=100000; k.budget_income=95000; k.actual_expenses=60000
k.actual_noi=40000; k.budget_noi=37000; k.eco_occ_pct=0.94; k.gpr=100000

ar = [
    ARAgingRow('Test Prop','PM Co','f.xlsx','Tenant Rent',2024,3,50000,40000,2000,1500,1000,500,-200),
    ARAgingRow('Test Prop','PM Co','f.xlsx','Tenant Rent',2025,3,55000,44000,2200,1800,1200,600,-250),
]
path = os.path.join(tempfile.gettempdir(), 'test_main_ar.xlsx')
build_main_workbook([k], 'Test Portfolio', path, ar_rows=ar)
import openpyxl
wb = openpyxl.load_workbook(path, read_only=True)
assert 'AR Aging' in wb.sheetnames, f'Missing AR Aging tab. Tabs: {wb.sheetnames}'
print('OK — tabs:', wb.sheetnames)
"
```

Expected: `OK — tabs: ['Dashboard', 'Property Analysis', 'Property Monthly KPIs', 'AR Aging']`

- [ ] **Step 3: Smoke test with no AR rows (empty state)**

```powershell
python -c "
import os, tempfile
from app.models import PropertyPeriodKPIs
from app.exporter.main_workbook import build_main_workbook

k = PropertyPeriodKPIs('Test Prop','PM Co',2024,1,'Jan')
k.actual_noi=40000
path = os.path.join(tempfile.gettempdir(), 'test_main_noar.xlsx')
build_main_workbook([k], 'Test Portfolio', path)
import openpyxl
wb = openpyxl.load_workbook(path, read_only=True)
assert 'AR Aging' in wb.sheetnames
print('OK — empty AR Aging tab present')
"
```

Expected: `OK — empty AR Aging tab present`

- [ ] **Step 4: Commit**

```powershell
git add app/exporter/main_workbook.py
git commit -m "feat: main workbook AR Aging tab (portfolio summary + property analysis blocks)"
```

---

## Task 9: Backup Workbook AR_Aging_Detail Tab

**Files:**
- Modify: `app/exporter/backup_workbook.py`

- [ ] **Step 1: Update `build_backup_workbook` signature and add `_build_ar_aging_detail`**

**a.** Update the `build_backup_workbook` function signature and body. Replace:

```python
def build_backup_workbook(
    mapped_rows: list[MappedRow],
    kpis: list[PropertyPeriodKPIs],
    source_index: list[SourceIndexEntry],
    mapping_entries: list[MappingEntry],
    quality_checks: list[QualityCheck],
    output_path: str,
    eco_occ_target: float = ECO_OCC_TARGET,
) -> str:
    """Builds backup workbook at output_path. Returns path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_raw_data(wb, mapped_rows)
    _build_source_index(wb, source_index)
    _build_assumptions_mapping(wb, mapping_entries)
    _build_budget_vs_actual(wb, kpis)
    _build_account_detail(wb, mapped_rows)
    _build_economic_occupancy(wb, kpis)
    _build_quality_checks(wb, quality_checks, kpis, eco_occ_target)

    wb.save(output_path)
    return output_path
```

With:

```python
def build_backup_workbook(
    mapped_rows: list[MappedRow],
    kpis: list[PropertyPeriodKPIs],
    source_index: list[SourceIndexEntry],
    mapping_entries: list[MappingEntry],
    quality_checks: list[QualityCheck],
    output_path: str,
    eco_occ_target: float = ECO_OCC_TARGET,
    ar_rows: list | None = None,
) -> str:
    """Builds backup workbook at output_path. Returns path."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_raw_data(wb, mapped_rows)
    _build_source_index(wb, source_index)
    _build_assumptions_mapping(wb, mapping_entries)
    _build_budget_vs_actual(wb, kpis)
    _build_account_detail(wb, mapped_rows)
    _build_economic_occupancy(wb, kpis)
    _build_quality_checks(wb, quality_checks, kpis, eco_occ_target)
    _build_ar_aging_detail(wb, ar_rows)

    wb.save(output_path)
    return output_path
```

**b.** Append the `_build_ar_aging_detail` function at the end of `app/exporter/backup_workbook.py` (after `_month_name`):

```python
def _build_ar_aging_detail(wb, ar_rows):
    """Build the 'AR_Aging_Detail' tab with one row per ARAgingRow."""
    ws = wb.create_sheet("AR_Aging_Detail")
    headers = [
        "Property", "PM", "Source File", "Receivable Type",
        "Year", "Month", "Period",
        "Charge Amount", "Current Owed", "0-30 Owed",
        "31-60 Owed", "61-90 Owed", "Over 90 Owed",
        "Pre-payments", "Total Overdue", "% >30 Days",
    ]
    _write_header(ws, headers, 1)

    if not ar_rows:
        ws.cell(2, 1, "No AR Aging data uploaded.")
        return

    # Sort: Receivable Type → Property Name → Year → Month
    sorted_rows = sorted(ar_rows,
                         key=lambda r: (r.receivable_type, r.property_name, r.year, r.month))

    for i, r in enumerate(sorted_rows, 2):
        total_overdue = r.owed_31_60 + r.owed_61_90 + r.owed_over_90
        pct_overdue   = total_overdue / r.charge_amount if r.charge_amount > 0 else None

        ws.cell(i,  1, r.property_name)
        ws.cell(i,  2, r.pm_name)
        ws.cell(i,  3, r.source_file)
        ws.cell(i,  4, r.receivable_type)
        ws.cell(i,  5, r.year)
        ws.cell(i,  6, r.month)
        ws.cell(i,  7, _month_name(r.month))
        ws.cell(i,  8, r.charge_amount);   ws.cell(i,  8).number_format = CURRENCY_FMT
        ws.cell(i,  9, r.current_owed);    ws.cell(i,  9).number_format = CURRENCY_FMT
        ws.cell(i, 10, r.owed_0_30);       ws.cell(i, 10).number_format = CURRENCY_FMT
        ws.cell(i, 11, r.owed_31_60);      ws.cell(i, 11).number_format = CURRENCY_FMT
        ws.cell(i, 12, r.owed_61_90);      ws.cell(i, 12).number_format = CURRENCY_FMT
        ws.cell(i, 13, r.owed_over_90);    ws.cell(i, 13).number_format = CURRENCY_FMT
        ws.cell(i, 14, r.prepayments);     ws.cell(i, 14).number_format = CURRENCY_FMT
        ws.cell(i, 15, total_overdue);     ws.cell(i, 15).number_format = CURRENCY_FMT
        ws.cell(i, 16, pct_overdue);       ws.cell(i, 16).number_format = PCT_FMT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
```

- [ ] **Step 2: Smoke test the backup workbook builder**

```powershell
python -c "
import os, tempfile
from app.models import ARAgingRow
from app.exporter.backup_workbook import build_backup_workbook

ar = [ARAgingRow('Test Prop','PM','f.xlsx','Tenant Rent',2024,3,50000,40000,2000,1500,1000,500,-200)]
path = os.path.join(tempfile.gettempdir(), 'test_backup_ar.xlsx')
build_backup_workbook([], [], [], [], [], path, ar_rows=ar)
import openpyxl
wb = openpyxl.load_workbook(path, read_only=True)
assert 'AR_Aging_Detail' in wb.sheetnames, f'Missing tab. Tabs: {wb.sheetnames}'
print('OK — tabs:', wb.sheetnames)
"
```

Expected: `OK — tabs: ['Raw_Data', 'Source_Index', 'Assumptions_Mapping', 'Budget_vs_Actual', 'Account_Detail', 'Economic_Occupancy', 'Quality_Checks', 'AR_Aging_Detail']`

- [ ] **Step 3: Commit**

```powershell
git add app/exporter/backup_workbook.py
git commit -m "feat: backup workbook AR_Aging_Detail tab (full raw AR data)"
```

---

## Task 10: Validator Update + Full Test Suite

**Files:**
- Modify: `app/exporter/validator.py`

- [ ] **Step 1: Update expected tab sets in `app/exporter/validator.py`**

Replace the two constants:

```python
_MAIN_TABS    = {"Dashboard", "Property Analysis", "Property Monthly KPIs"}
_BACKUP_TABS  = {"Raw_Data", "Source_Index", "Assumptions_Mapping",
                 "Budget_vs_Actual", "Account_Detail", "Economic_Occupancy", "Quality_Checks"}
```

With:

```python
_MAIN_TABS    = {"Dashboard", "Property Analysis", "Property Monthly KPIs", "AR Aging"}
_BACKUP_TABS  = {"Raw_Data", "Source_Index", "Assumptions_Mapping",
                 "Budget_vs_Actual", "Account_Detail", "Economic_Occupancy",
                 "Quality_Checks", "AR_Aging_Detail"}
```

- [ ] **Step 2: Run the full test suite**

```powershell
cd C:\Users\erwin\Desktop\portfolio-analyzer
pytest tests/ -v
```

Expected: All tests PASS (including the 15+ new AR aging parser tests and all pre-existing tests).

- [ ] **Step 3: End-to-end smoke test — workbooks build and pass validation**

```powershell
python -c "
import os, tempfile
from app.models import PropertyPeriodKPIs, ARAgingRow
from app.exporter.main_workbook import build_main_workbook
from app.exporter.backup_workbook import build_backup_workbook
from app.exporter.validator import validate_both_workbooks

k = PropertyPeriodKPIs('Prop A','PM Co',2024,1,'Jan')
k.actual_income=100000; k.actual_expenses=60000; k.actual_noi=40000
k.eco_occ_pct=0.94; k.gpr=100000; k.vacancy=3000; k.concessions=1000; k.bad_debt=2000

ar = [
    ARAgingRow('Prop A','PM Co','f.xlsx','Tenant Rent',2024,3,50000,40000,2000,1500,1000,500,-200),
    ARAgingRow('Prop A','PM Co','f.xlsx','Subsidy',2024,3,30000,25000,1000,800,500,300,0),
    ARAgingRow('Prop A','PM Co','f.xlsx','Tenant Rent',2025,3,55000,44000,2200,1800,1200,600,-250),
]

tmp = tempfile.gettempdir()
main_path   = os.path.join(tmp, 'smoke_main.xlsx')
backup_path = os.path.join(tmp, 'smoke_backup.xlsx')

build_main_workbook([k], 'Smoke Test', main_path, ar_rows=ar)
build_backup_workbook([], [k], [], [], [], backup_path, ar_rows=ar)

checks = validate_both_workbooks(main_path, backup_path)
failures = [c for c in checks if not c.passed]
if failures:
    for f in failures:
        print('FAIL:', f.check_name, '|', f.detail)
else:
    print('All validation checks PASS')
    print('Tabs in main:  ', __import__(\"openpyxl\").load_workbook(main_path, read_only=True).sheetnames)
    print('Tabs in backup:', __import__(\"openpyxl\").load_workbook(backup_path, read_only=True).sheetnames)
"
```

Expected output:
```
All validation checks PASS
Tabs in main:   ['Dashboard', 'Property Analysis', 'Property Monthly KPIs', 'AR Aging']
Tabs in backup: ['Raw_Data', 'Source_Index', 'Assumptions_Mapping', 'Budget_vs_Actual', 'Account_Detail', 'Economic_Occupancy', 'Quality_Checks', 'AR_Aging_Detail']
```

- [ ] **Step 4: Commit**

```powershell
git add app/exporter/validator.py
git commit -m "feat: update validator tab sets for AR Aging and AR_Aging_Detail tabs"
```

- [ ] **Step 5: Final commit message summarizing the feature**

```powershell
git log --oneline -10
```

Verify the last 10 commits show the complete AR Aging feature chain.

---

## Verification Checklist

After all tasks complete, verify end-to-end:

- [ ] `pytest tests/ -v` — all tests green
- [ ] `python app.py` — app starts without errors
- [ ] Upload page shows "AR Aging Reports (Optional)" drop zone between Physical Occupancy and Analysis Settings sections
- [ ] Upload AR Aging files → dashboard shows "Portfolio AR Summary — Tenant Rent" and "Portfolio AR Summary — Subsidy" sections with period columns
- [ ] YoY Δ column appears only for periods that have a prior-year same-month match
- [ ] YoY Δ increases in Current Owed or % >30 Days shown in red; decreases in green
- [ ] "Property AR Analysis" table below sub-tables, sorted by TR Current Owed descending
- [ ] Property detail page shows "AR Aging" section with Tenant Rent and Subsidy sub-tables per period
- [ ] Download ZIP → main workbook has "AR Aging" as 4th tab with all 4 blocks
- [ ] Download ZIP → backup workbook has "AR_Aging_Detail" as 8th tab with one row per ARAgingRow
- [ ] Running without any AR files → AR Aging tab still present (empty state message); no errors; validation passes
- [ ] Existing runs (without `ar_aging.json`) reload correctly — dashboard shows no AR section, no crash
