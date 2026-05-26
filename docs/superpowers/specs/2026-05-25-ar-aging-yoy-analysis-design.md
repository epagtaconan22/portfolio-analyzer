# AR Aging YoY Trend Analysis — Design Spec

_Date: 2026-05-25 | Author: Erwin @ Affirmed Housing Group_

---

## Context

Affirmed Housing Group receives AR Aging reports from each PM company at the end of each reporting period. These are Yardi-generated "Affordable Aging Detail" spreadsheets, delivered separately for Tenant Rent and Subsidy receivables. Currently there is no tool to aggregate these across properties, compare them year-over-year, or display them alongside the existing NOI and occupancy KPIs.

This spec adds an AR Aging YoY trend analysis layer to the portfolio-analyzer application, covering:
- A parser for Yardi AR aging export files
- Upload UI support for AR aging files
- Portfolio and property-level summary views (three headline metrics) in the web dashboard
- A dedicated AR Aging tab in the main Excel workbook
- Full raw AR detail in the backup workbook

---

## File Format — Yardi AR Aging Export

All AR aging files (Tenant Rent and Subsidy) share an identical 9-column layout on a single sheet named `Report1`:

| Row | Content |
|---|---|
| 1 | Report title (e.g., `"Affordable Aging Detail"`) |
| 2 | `"Property: Affirmed Property List (affirmed)"` |
| 3 | `"Post To(MM/YY): 03/2024"` — report period (month/year) |
| 4–5 | Split two-row column header (combined: Property Name, Charge Amount, Current Owed, 0-30 Owed, 31-60 Owed, 61-90 Owed, Over 90 Owed, Pre-payments, Suspense) |
| 6 … N-2 | Property data rows |
| N-1 | Blank row |
| N | `"Grand Total"` summary row |

**Column mapping (0-indexed):**

| Index | Name | Notes |
|---|---|---|
| 0 | Property Name | Format: `"Property Name (code)"` e.g. `"Alora Family (alora)"` |
| 1 | Charge Amount | Total charges outstanding |
| 2 | Current Owed | Current period balance |
| 3 | 0-30 Owed | 0–30 day aging bucket |
| 4 | 31-60 Owed | 31–60 day aging bucket |
| 5 | 61-90 Owed | 61–90 day aging bucket |
| 6 | Over 90 Owed | 90+ day aging bucket |
| 7 | Pre-payments | Credit balances (always negative or zero) |
| 8 | Suspense | Always 0 — ignored |

---

## File Naming Convention

Expected pattern: `PMC_AR Aging_[Type]_MM_YYYY.xlsx`

- `PMC` — PM company prefix (e.g., `Solari`, `ConAm`)
- `[Type]` — one of: `Tenant Rent`, `Tenant Receivable`, `Subsidy`, `Subsidy Receivable`
- `MM_YYYY` — two-digit month and four-digit year

The parser normalizes type variants:
- `"Tenant Rent"` and `"Tenant Receivable"` → canonical `"Tenant Rent"`
- `"Subsidy"` and `"Subsidy Receivable"` → canonical `"Subsidy"`

If the filename does not match the pattern, the parser falls back to parsing row 3 of the sheet for the period and uses the filename stem as the type hint.

---

## Data Model

### `ARAgingRow` (new dataclass in `app/models.py`)

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
        """31-60 + 61-90 + over-90 — amounts past the current + 0-30 buckets."""
        return self.owed_31_60 + self.owed_61_90 + self.owed_over_90

    @property
    def pct_overdue(self) -> Optional[float]:
        """% of charge_amount that is >30 days past due."""
        if self.charge_amount and self.charge_amount > 0:
            return self.total_overdue / self.charge_amount
        return None
```

---

## Parser — `app/parser/ar_aging.py`

**Function signature:**
```python
def parse_ar_aging_reports(file_paths: list[str]) -> list[ARAgingRow]:
```

**Logic:**
1. For each file, extract `pm_name` from the filename prefix (text before the first `_`).
2. Extract `receivable_type` from the filename segment between the second and third `_` separators; normalize to `"Tenant Rent"` or `"Subsidy"`.
3. Extract `year` and `month` from the last two `_`-separated segments before `.xlsx`.
4. Fall back: if filename doesn't match, parse row 3 of the sheet for period; infer type from filename stem keywords.
5. Open the workbook (`data_only=True`), read `Report1`.
6. Data rows start at row index 5 (0-based). Stop at the first row where column A is `None` or the string starts with `"Grand Total"`.
7. For each data row: extract property name (strip the `(code)` suffix), all 7 numeric columns (treat `None` as `0.0`).
8. Apply `PROPERTY_NAME_MAP` to `property_name`.
9. Skip rows where `property_name` is blank after stripping.
10. Return list of `ARAgingRow` records.

**Property name extraction:** strip the code in parentheses at the end:
```python
import re
name = re.sub(r'\s*\([^)]+\)\s*$', '', raw_name).strip()
```

---

## Upload UI — `app/templates/upload.html`

Add a new optional drop zone section **"AR Aging Reports (Optional)"** between the Physical Occupancy section and the Analysis Settings section.

- Input field name: `ar_aging_files` (multiple)
- Accepts: `.xlsx`, `.xls`
- Hint text: `"Files named PMC_AR Aging_Tenant Rent_MM_YYYY.xlsx or PMC_AR Aging_Subsidy_MM_YYYY.xlsx. Multiple files accepted — one per period per type per PM company."`

---

## Pipeline Integration — `app/routes/upload.py`

After the occupancy parsing step:

```python
ar_rows = []
ar_files = request.files.getlist("ar_aging_files")
for ar_file in ar_files:
    if ar_file and ar_file.filename:
        ar_path = os.path.join("uploads", secure_filename(ar_file.filename))
        ar_file.save(ar_path)
        ar_rows.extend(parse_ar_aging_reports([ar_path]))

# Apply PROPERTY_NAME_MAP (parser already does this, but apply again for safety)
for _row in ar_rows:
    _row.property_name = PROPERTY_NAME_MAP.get(_row.property_name, _row.property_name)
```

Pass `ar_rows` to `build_main_workbook`, `build_backup_workbook`, and `save_run`.

---

## Storage — `app/storage/runs.py`

`save_run` gains an `ar_rows: list[ARAgingRow]` parameter and writes `ar_aging.json`.
`load_run` returns a dict with an `"ar_aging"` key (list of dicts).

Two new metadata fields:
- `ar_tenant_rent_periods`: sorted list of `"MM-YYYY"` strings from Tenant Rent files
- `ar_subsidy_periods`: sorted list of `"MM-YYYY"` strings from Subsidy files

---

## Dashboard — `app/templates/dashboard.html` + `app/routes/results.py`

### New "AR Aging" section on the dashboard

Placed below the per-property table. Two sub-tables rendered sequentially:

**Sub-table: Portfolio AR Summary — Tenant Rent (N properties)**  
**Sub-table: Portfolio AR Summary — Subsidy (N properties)**

Each sub-table is a transposed table (rows = metrics, columns = periods):

| Metric | Mar-2024 | Mar-2025 | YoY Δ | Jun-2025 | Sep-2025 | Dec-2025 | Mar-2026 | YoY Δ |
|---|---|---|---|---|---|---|---|---|
| Current Owed | $X | $X | +$X | $X | $X | $X | $X | +$X |
| Pre-payments | | | | | | | | |
| % >30 Days | | | | | | | | |

- Columns are all uploaded periods in ascending chronological order.
- A **YoY Δ column** is inserted immediately after each period for which the same calendar month exists one year prior (e.g., a YoY Δ column after Mar-2025 because Mar-2024 was uploaded; after Mar-2026 because Mar-2025 was uploaded). No YoY Δ column appears for periods with no prior-year match (e.g., Jun-2025 if Jun-2024 was not uploaded).
- YoY Δ for Current Owed and Pre-payments is shown in $ (positive = increase in balance). YoY Δ for % >30 Days is shown in percentage points.
- Color coding: YoY Δ increases in Current Owed or % >30 Days are colored red (unfavorable); decreases are green.
- Property count N in the sub-table header = distinct property names with data for that receivable type.

### New per-property AR table

Below the sub-tables, a flat property-level table for the **latest AR period** (most recent by date across all uploaded AR files):

| Property | PM | TR Current Owed | TR Pre-payments | TR % >30 | TR YoY Δ % >30 | Sub Current Owed | Sub Pre-payments | Sub % >30 | Sub YoY Δ % >30 |
|---|---|---|---|---|---|---|---|---|---|

- `TR` = Tenant Rent, `Sub` = Subsidy.
- If a property has no data for one type, show `—`.
- Sort by Tenant Rent Current Owed descending (highest AR exposure first).

### Results route additions (`app/routes/results.py`)

New helper `_agg_ar(ar_rows: list[dict], receivable_type: str, year: int, month: int) -> dict`:
- Filters to matching type/period, sums `current_owed`, `prepayments`, `total_overdue` (= `owed_31_60 + owed_61_90 + owed_over_90`), `charge_amount`.
- Returns `{current_owed, prepayments, pct_overdue, property_count}`.

New helper `_ar_yoy_delta(curr: dict, prev: dict) -> dict`:
- Returns `{current_owed_delta, prepayments_delta, pct_overdue_delta}`.

Pass to template:
- `ar_summary`: `{receivable_type → {period_label → agg_dict}}`
- `ar_periods`: sorted list of `(year, month)` tuples with their labels
- `ar_yoy_map`: `{receivable_type → {period_label → delta_dict}}` (only where prior-year same-month exists)
- `ar_prop_rows`: list of per-property dicts for the latest period

---

## Property Detail Page — `app/templates/property_detail.html`

New **"AR Aging"** section after the existing Monthly KPIs table.

Two sub-tables per property (Tenant Rent, Subsidy — only shown if data exists):

| Period | Charge Amt | Current | 0–30 | 31–60 | 61–90 | Over 90 | Pre-payments | % >30 Days |
|---|---|---|---|---|---|---|---|---|

Rows = all uploaded periods for this property, sorted chronologically. This is the full 7-column breakdown (charge amount + all buckets) since it's a detail page for a single property.

---

## Main Workbook — `app/exporter/main_workbook.py`

### New "AR Aging" tab (4th tab)

**Block 1: Portfolio AR Summary — Tenant Rent** (header row shows property count)
- Rows: Current Owed, Pre-payments, % >30 Days
- Columns: all periods chronologically + YoY Δ columns where prior-year same-month exists
- Currency format for Current Owed and Pre-payments; percentage format for % >30 Days
- YoY Δ columns: currency for dollar metrics, percentage for % >30 Days
- Conditional formatting: YoY Δ increases in $ balances or % >30 = red fill; decreases = green fill

**Block 2: Portfolio AR Summary — Subsidy** (same structure, separated by 2 blank rows)

**Block 3: Property AR Analysis — Tenant Rent** (header: "As of [latest period]")
One row per property. Columns:

| Property | PM | Current Owed | Pre-payments | % >30 Days | YoY $ Δ (Current) | YoY Δ (% >30) |

**Block 4: Property AR Analysis — Subsidy** (same structure)

KPI column headers include Excel cell comments with formula definitions (consistent with existing workbook style).

Frozen panes at row below each block header. Autofilter on property analysis blocks.

---

## Backup Workbook — `app/exporter/backup_workbook.py`

### New "AR_Aging_Detail" tab (8th tab)

One row per `ARAgingRow`. All fields plus computed columns:

| Property | PM | Source File | Receivable Type | Year | Month | Period | Charge Amount | Current Owed | 0–30 | 31–60 | 61–90 | Over 90 | Pre-payments | Total Overdue | % >30 Days |

- Autofilter on header row.
- Frozen panes at row 2.
- Currency format for all dollar columns; percentage for % >30 Days.
- Sorted by: Receivable Type, Property Name, Year, Month.

---

## Validator Update — `app/exporter/validator.py`

Update expected tab sets:
```python
_MAIN_TABS   = {"Dashboard", "Property Analysis", "Property Monthly KPIs", "AR Aging"}
_BACKUP_TABS = {"Raw_Data", "Source_Index", "Assumptions_Mapping",
                "Budget_vs_Actual", "Account_Detail", "Economic_Occupancy",
                "Quality_Checks", "AR_Aging_Detail"}
```

---

## Tests

New file: `tests/test_ar_aging_parser.py`

Test cases:
- Parse a well-formed Tenant Rent file → correct property names, year, month, all 7 numeric fields
- Parse a Subsidy file → `receivable_type == "Subsidy"`
- Type normalization: `"Tenant Receivable"` filename → `"Tenant Rent"`
- Property name stripping: `"Alora Family (alora)"` → `"Alora Family"` (before map) → canonical name after map
- Grand Total row excluded
- Blank row excluded
- `pct_overdue` computed correctly: `(owed_31_60 + owed_61_90 + owed_over_90) / charge_amount`
- Fixture: synthetic `tmp_path` workbook matching Yardi layout

---

## File Change Summary

| File | Change |
|---|---|
| `app/models.py` | Add `ARAgingRow` dataclass |
| `app/parser/ar_aging.py` | New — AR aging parser |
| `tests/test_ar_aging_parser.py` | New — parser tests |
| `app/templates/upload.html` | Add AR aging drop zone |
| `app/routes/upload.py` | Parse AR files, pass to exporters and save_run |
| `app/storage/runs.py` | Add `ar_aging.json` save/load; new metadata fields |
| `app/routes/results.py` | Add `_agg_ar`, `_ar_yoy_delta`; pass AR context to template |
| `app/templates/dashboard.html` | Add AR Aging section (portfolio sub-tables + property table) |
| `app/templates/property_detail.html` | Add AR Aging detail section |
| `app/exporter/main_workbook.py` | Add AR Aging tab (4 blocks) |
| `app/exporter/backup_workbook.py` | Add AR_Aging_Detail tab |
| `app/exporter/validator.py` | Update expected tab sets |
