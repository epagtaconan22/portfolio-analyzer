# Streamlit Port Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Flask/Jinja2 UI layer with Streamlit while keeping all backend logic (parsers, calculators, exporters, storage) completely unchanged.

**Architecture:** Extract reusable pure-Python helpers from the Flask route files into a new `app/ui/` module (formatting, aggregation, KPI definitions, pipeline runner, projection), then write four Streamlit pages (`pages/upload.py`, `pages/dashboard.py`, `pages/property_detail.py`, `pages/history.py`) that import from `app/ui/`. The Flask layer (`app/routes/`, `app/templates/`, `app/__init__.py`, `app.py`, `static/style.css`) is deleted only in the final task after all tests pass.

**Tech Stack:** Python 3.11+, Streamlit ≥ 1.36, pandas 2.x (Styler for color-coded tables), openpyxl (unchanged), pytest (unchanged)

---

## File Map

### New files (create)
| File | Purpose |
|---|---|
| `streamlit_app.py` | Entry point — `st.navigation()` wiring |
| `.streamlit/config.toml` | Navy theme |
| `app/ui/__init__.py` | Empty package marker |
| `app/ui/formatting.py` | `fmt_currency()`, `fmt_pct()` replacing Jinja2 filters |
| `app/ui/aggregation.py` | `agg_kpis()`, `agg_ar()`, quarter helpers — extracted from `app/routes/results.py` |
| `app/ui/kpi_definitions.py` | `SUMMARY_KPI_DEFINITIONS`, `KPI_TOOLTIPS`, YoY constants — extracted from `app/routes/results.py` |
| `app/ui/projection.py` | `compute_prop_projection()` — extracted from `app/routes/property_detail.py` |
| `app/ui/pipeline.py` | `run_analysis_pipeline()` — business logic extracted from `app/routes/upload.py` |
| `pages/upload.py` | Streamlit upload form |
| `pages/dashboard.py` | Streamlit dashboard |
| `pages/property_detail.py` | Streamlit property detail |
| `pages/history.py` | Streamlit history |
| `tests/test_pipeline.py` | Replaces `tests/test_routes.py` |

### Modified files
| File | Change |
|---|---|
| `requirements.txt` | Add `streamlit>=1.36`, remove `flask>=3.0`, `pytest-flask>=1.3` |

### Deleted (Task 12 only — after all tests pass)
`app/routes/`, `app/templates/`, `app/__init__.py`, `app.py`, `static/style.css`, `tests/test_routes.py`

### Untouched (do not modify)
`app/parser/`, `app/mapper/`, `app/calculator/`, `app/exporter/`, `app/storage/`, `app/models.py`, `config.py`, all other test files

---

## Task 1: Dependencies, Entry Point, and Theme

**Files:**
- Modify: `requirements.txt`
- Create: `streamlit_app.py`
- Create: `.streamlit/config.toml`
- Create: `app/ui/__init__.py`
- Create: `pages/` directory

- [ ] **Step 1: Update requirements.txt**

Replace the contents of `requirements.txt` with:

```
streamlit>=1.36
openpyxl>=3.1
pandas>=2.0
xlrd>=2.0
pytest>=8.0
```

Note: `flask`, `jinja2`, and `pytest-flask` are removed. `werkzeug` (used in upload.py for `secure_filename`) will be replaced with a direct `re.sub` sanitizer in `app/ui/pipeline.py`.

- [ ] **Step 2: Install the updated dependencies**

```powershell
cd C:\Users\erwin\Desktop\portfolio-analyzer
pip install -r requirements.txt
```

Expected: Streamlit installs successfully, no errors.

- [ ] **Step 3: Create `.streamlit/config.toml`**

```toml
[theme]
primaryColor = "#1F4E79"
backgroundColor = "#F4F6F9"
secondaryBackgroundColor = "#FFFFFF"
textColor = "#1A1A2E"
font = "sans serif"

[server]
headless = false
port = 8501
```

- [ ] **Step 4: Create `app/ui/__init__.py`**

```python
```

(Empty file — just the package marker.)

- [ ] **Step 5: Create the `pages/` directory with placeholder files**

Create `pages/upload.py`:
```python
import streamlit as st
st.title("Upload — coming soon")
```

Create `pages/dashboard.py`:
```python
import streamlit as st
st.title("Dashboard — coming soon")
```

Create `pages/property_detail.py`:
```python
import streamlit as st
st.title("Property Detail — coming soon")
```

Create `pages/history.py`:
```python
import streamlit as st
st.title("History — coming soon")
```

- [ ] **Step 6: Create `streamlit_app.py`**

```python
import streamlit as st

st.set_page_config(
    page_title="Portfolio Analyzer",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

upload_page  = st.Page("pages/upload.py",          title="New Analysis",    icon="📤", default=True)
dash_page    = st.Page("pages/dashboard.py",        title="Dashboard",       icon="📊")
detail_page  = st.Page("pages/property_detail.py",  title="Property Detail", icon="🏢")
history_page = st.Page("pages/history.py",          title="History",         icon="🕐")

nav = st.navigation(
    [upload_page, dash_page, detail_page, history_page],
    position="sidebar",
)
nav.run()
```

- [ ] **Step 7: Verify the app starts**

```powershell
streamlit run streamlit_app.py
```

Expected: Browser opens at `http://localhost:8501`. The sidebar shows four pages. "New Analysis" (placeholder) is selected by default. No errors in the terminal.

Stop the server with Ctrl-C.

- [ ] **Step 8: Run existing tests to confirm nothing broke**

```powershell
pytest tests/ -q --ignore=tests/test_routes.py
```

Expected: All tests except the Flask route tests pass. (We'll replace `test_routes.py` in Task 11.)

- [ ] **Step 9: Commit**

```powershell
git add requirements.txt .streamlit/config.toml streamlit_app.py app/ui/__init__.py pages/
git commit -m "feat: scaffold Streamlit entry point, pages, theme"
```

---

## Task 2: Shared Formatting Helpers

**Files:**
- Create: `app/ui/formatting.py`
- Create: `tests/test_formatting.py`

These two functions replace the Jinja2 `currency` and `pct` template filters throughout all Streamlit pages.

- [ ] **Step 1: Write failing tests**

Create `tests/test_formatting.py`:

```python
from app.ui.formatting import fmt_currency, fmt_pct


def test_fmt_currency_positive():
    assert fmt_currency(1234567) == "$1,234,567"


def test_fmt_currency_negative():
    assert fmt_currency(-50000) == "($50,000)"


def test_fmt_currency_zero():
    assert fmt_currency(0) == "$0"


def test_fmt_currency_none():
    assert fmt_currency(None) == "—"


def test_fmt_currency_float():
    assert fmt_currency(99999.99) == "$100,000"


def test_fmt_pct_positive():
    assert fmt_pct(0.954) == "95.4%"


def test_fmt_pct_negative():
    assert fmt_pct(-0.032) == "-3.2%"


def test_fmt_pct_zero():
    assert fmt_pct(0) == "0.0%"


def test_fmt_pct_none():
    assert fmt_pct(None) == "—"
```

- [ ] **Step 2: Run tests to confirm they fail**

```powershell
pytest tests/test_formatting.py -v
```

Expected: All FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Create `app/ui/formatting.py`**

```python
"""Formatting helpers — replaces Jinja2 currency and pct template filters."""

from typing import Optional


def fmt_currency(value: Optional[float]) -> str:
    """Format a dollar amount: $1,234,567 or ($50,000) for negatives. '—' for None."""
    if value is None:
        return "—"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "—"
    if v < 0:
        return f"(${abs(v):,.0f})"
    return f"${v:,.0f}"


def fmt_pct(value: Optional[float]) -> str:
    """Format a decimal ratio as a percentage: 0.954 → '95.4%'. '—' for None."""
    if value is None:
        return "—"
    try:
        return f"{float(value):.1%}"
    except (TypeError, ValueError):
        return "—"
```

- [ ] **Step 4: Run tests to confirm they pass**

```powershell
pytest tests/test_formatting.py -v
```

Expected: All 10 PASS.

- [ ] **Step 5: Commit**

```powershell
git add app/ui/formatting.py tests/test_formatting.py
git commit -m "feat: shared formatting helpers (fmt_currency, fmt_pct)"
```

---

## Task 3: Extract Aggregation Helpers

**Files:**
- Create: `app/ui/aggregation.py`
- Create: `tests/test_aggregation.py`

These functions are currently inside `app/routes/results.py`. Extracting them makes them importable by Streamlit pages without importing Flask.

- [ ] **Step 1: Write failing tests**

Create `tests/test_aggregation.py`:

```python
import pytest
from app.ui.aggregation import agg_kpis, month_to_quarter, quarter_label, quarter_months


def _kpi(income, expenses, gpr=None, vacancy=0, year=2025, month=1):
    noi = income - expenses if income is not None and expenses is not None else None
    net_coll = (gpr - vacancy) if gpr is not None else None
    return {
        "actual_income": income,
        "actual_expenses": expenses,
        "actual_noi": noi,
        "budget_income": income * 0.95 if income else None,
        "budget_expenses": expenses * 1.05 if expenses else None,
        "budget_noi": None,
        "gpr": gpr,
        "vacancy": vacancy,
        "concessions": 0,
        "bad_debt": 0,
        "net_collectible": net_coll,
        "eco_occ_pct": (net_coll / gpr) if (net_coll and gpr) else None,
        "budget_eco_occ_pct": None,
        "physical_occ_pct": None,
        "occupied_units": None,
        "total_units": 100,
        "income_per_unit": None,
        "expense_per_unit": None,
        "noi_per_unit": None,
        "year": year,
        "month": month,
    }


def test_agg_kpis_sums_income():
    kpis = [_kpi(10000, 4000), _kpi(8000, 3000)]
    result = agg_kpis(kpis)
    assert result["actual_income"] == pytest.approx(18000)


def test_agg_kpis_computes_noi():
    kpis = [_kpi(10000, 4000)]
    result = agg_kpis(kpis)
    assert result["actual_noi"] == pytest.approx(6000)


def test_agg_kpis_noi_variance():
    kpis = [_kpi(10000, 4000)]
    result = agg_kpis(kpis)
    assert result["noi_variance"] is not None


def test_agg_kpis_empty_returns_all_none():
    result = agg_kpis([])
    assert result["actual_income"] is None
    assert result["actual_noi"] is None


def test_month_to_quarter():
    assert month_to_quarter(1) == 1
    assert month_to_quarter(3) == 1
    assert month_to_quarter(4) == 2
    assert month_to_quarter(12) == 4


def test_quarter_label():
    assert quarter_label(2025, 1) == "Q1 - 2025"
    assert quarter_label(2026, 4) == "Q4 - 2026"


def test_quarter_months():
    assert quarter_months(1) == {1, 2, 3}
    assert quarter_months(4) == {10, 11, 12}
```

- [ ] **Step 2: Run tests to confirm they fail**

```powershell
pytest tests/test_aggregation.py -v
```

Expected: All FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Create `app/ui/aggregation.py`**

Copy the four helper functions verbatim from `app/routes/results.py` (lines 110–279). The only change is removing the leading underscore from names that were module-private:

```python
"""KPI and AR aggregation helpers — extracted from routes/results.py."""

from typing import Optional


# ── Quarter helpers ────────────────────────────────────────────────────────────

def month_to_quarter(month: int) -> int:
    return (month - 1) // 3 + 1


def quarter_label(year: int, quarter: int) -> str:
    return f"Q{quarter} - {year}"


def quarter_months(quarter: int) -> set[int]:
    base = (quarter - 1) * 3 + 1
    return {base, base + 1, base + 2}


# ── AR Aging helpers ───────────────────────────────────────────────────────────

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}


def ar_period_label(year: int, month: int) -> str:
    return f"{_MONTH_ABBR.get(month, str(month))}-{year}"


def agg_ar(ar_rows: list[dict], receivable_type: str,
           year: int, month: int) -> dict | None:
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


def agg_ar_for_prop(ar_rows: list[dict], property_name: str,
                    receivable_type: str, year: int, month: int) -> dict | None:
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
        "pct_overdue":  (over_60 / charge) if charge > 0 else 0.0,
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


def pct_delta(curr: dict | None, prev: dict | None) -> float | None:
    if curr and prev:
        c = curr.get("pct_overdue")
        p = prev.get("pct_overdue")
        if c is not None and p is not None:
            return c - p
    return None


# ── KPI aggregation ────────────────────────────────────────────────────────────

def agg_kpis(kpi_dicts: list[dict]) -> dict:
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
```

- [ ] **Step 4: Run tests to confirm they pass**

```powershell
pytest tests/test_aggregation.py -v
```

Expected: All 8 PASS.

- [ ] **Step 5: Commit**

```powershell
git add app/ui/aggregation.py tests/test_aggregation.py
git commit -m "feat: extract aggregation helpers to app/ui/aggregation"
```

---

## Task 4: Extract KPI Definitions

**Files:**
- Create: `app/ui/kpi_definitions.py`

This is a pure data extraction — no logic, no tests needed. All five constants currently live in `app/routes/results.py` (lines 36–105).

- [ ] **Step 1: Create `app/ui/kpi_definitions.py`**

```python
"""KPI display definitions — extracted from routes/results.py."""

# Tooltip text for KPI labels
KPI_TOOLTIPS: dict[str, str] = {
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
    "NOI Variance %":     "NOI Variance / |Budget NOI|. Absolute denominator handles sign flips",
    "GPR":                "Gross Potential Rent — total scheduled rent before any deductions",
    "Vacancy":            "Vacancy loss — rent foregone from unoccupied units",
    "Concessions":        "Move-in specials and rent concessions",
    "Bad Debt":           "Collection losses and write-offs",
    "Net Collectible":    "GPR − Vacancy − Concessions − Bad Debt",
    "Eco Occ %":          "Economic Occupancy % = Net Collectible / GPR",
    "Budget Eco Occ %":   "Budget Economic Occupancy % = Budget Net Collectible / Budget GPR",
    "Eco Occ Variance":   "Actual Eco Occ % − Budget Eco Occ %",
    "Physical Occ %":     "Physical Occ % = Occupied Units / Total Units",
    "Leakage Gap":        "Physical Occ % − Economic Occ %. Positive = units occupied but rent not fully collected",
    "Income/Unit":        "Actual Income / Total Units (from Physical Occupancy Report)",
    "Expense/Unit":       "Actual Expenses / Total Units",
    "NOI/Unit":           "Actual NOI / Total Units",
}

# (label, key, fmt, favorable_positive, group_id, is_group_header)
# None entries are visual separators (blank rows in the table).
SUMMARY_KPI_DEFINITIONS = [
    ("Actual Income",      "actual_income",        "currency", None,  "group_income",   True),
    ("Budget Income",      "budget_income",        "currency", None,  "group_income",   False),
    ("Income Variance",    "income_variance",      "currency", True,  "group_income",   False),
    None,
    ("Actual Expenses",    "actual_expenses",      "currency", None,  "group_expenses", True),
    ("Budget Expenses",    "budget_expenses",      "currency", None,  "group_expenses", False),
    ("Expense Variance",   "expense_variance",     "currency", False, "group_expenses", False),
    None,
    ("Actual NOI",         "actual_noi",           "currency", None,  "group_noi",      True),
    ("Budget NOI",         "budget_noi",           "currency", None,  "group_noi",      False),
    ("NOI Variance",       "noi_variance",         "currency", True,  "group_noi",      False),
    None,
    ("GPR",                "gpr",                  "currency", None,  "group_gpr",      True),
    ("Vacancy",            "vacancy",              "currency", None,  "group_gpr",      False),
    ("Concessions",        "concessions",          "currency", None,  "group_gpr",      False),
    ("Bad Debt",           "bad_debt",             "currency", None,  "group_gpr",      False),
    ("Net Collectible",    "net_collectible",      "currency", None,  "group_gpr",      False),
    None,
    ("Eco Occ %",          "eco_occ_pct",          "pct",      None,  "group_eco_occ",  True),
    ("Budget Eco Occ %",   "budget_eco_occ_pct",   "pct",      None,  "group_eco_occ",  False),
    ("Eco Occ Variance",   "eco_occ_variance",     "pct",      True,  "group_eco_occ",  False),
    ("Physical Occ %",     "physical_occ_pct",     "pct",      None,  None,             False),
    ("Leakage Gap",        "leakage_gap",          "pct",      False, None,             False),
    None,
    ("Income/Unit",        "income_per_unit",      "currency", None,  None,             False),
    ("Expense/Unit",       "expense_per_unit",     "currency", None,  None,             False),
    ("NOI/Unit",           "noi_per_unit",         "currency", None,  None,             False),
]

# Maps actual KPI key → the budget equivalent for YoY budget comparison
BUDGET_YOY_KEY: dict[str, str] = {
    "actual_income":        "budget_income",
    "budget_income":        "budget_income",
    "income_variance":      "income_variance",
    "income_variance_pct":  "income_variance_pct",
    "actual_expenses":      "budget_expenses",
    "budget_expenses":      "budget_expenses",
    "expense_variance":     "expense_variance",
    "expense_variance_pct": "expense_variance_pct",
    "actual_noi":           "budget_noi",
    "budget_noi":           "budget_noi",
    "noi_variance":         "noi_variance",
    "noi_variance_pct":     "noi_variance_pct",
    "eco_occ_pct":          "budget_eco_occ_pct",
}

YOY_CURRENCY_KEYS = frozenset({
    "actual_income", "budget_income", "income_variance",
    "actual_expenses", "budget_expenses", "expense_variance",
    "actual_noi", "budget_noi", "noi_variance",
})

YOY_FAVORABLE_IF_POSITIVE = frozenset({
    "actual_income", "budget_income",
    "actual_noi",    "budget_noi",    "noi_variance",
    "eco_occ_pct",
})

PCT_VARIANCE_THRESHOLD_KEYS = frozenset({
    "income_variance_pct", "expense_variance_pct",
    "noi_variance_pct",    "eco_occ_variance",
})
```

- [ ] **Step 2: Verify it imports cleanly**

```powershell
python -c "from app.ui.kpi_definitions import SUMMARY_KPI_DEFINITIONS, KPI_TOOLTIPS; print(len(SUMMARY_KPI_DEFINITIONS), 'definitions')"
```

Expected: `27 definitions` (27 entries including None separators).

- [ ] **Step 3: Commit**

```powershell
git add app/ui/kpi_definitions.py
git commit -m "feat: extract KPI display definitions to app/ui/kpi_definitions"
```

---

## Task 5: Extract Projection Helper

**Files:**
- Create: `app/ui/projection.py`
- Create: `tests/test_projection.py`

`compute_prop_projection()` is currently duplicated: one copy in `app/routes/property_detail.py` and one in `app/exporter/main_workbook.py`. The Streamlit property detail page will import from `app/ui/projection.py`.

- [ ] **Step 1: Write failing tests**

Create `tests/test_projection.py`:

```python
import pytest
from app.ui.projection import compute_prop_projection


def _kpi(month, actual=None, budget=None, year=2026):
    return {
        "year": year, "month": month,
        "actual_income": actual, "budget_income": budget,
        "actual_expenses": actual * 0.4 if actual else None,
        "budget_expenses": budget * 0.4 if budget else None,
        "actual_noi": actual * 0.6 if actual else None,
        "budget_noi": budget * 0.6 if budget else None,
    }


def test_returns_empty_when_no_kpis():
    label, data = compute_prop_projection([])
    assert label == ""
    assert data == {}


def test_uses_latest_year():
    kpis = [_kpi(1, 10000, 9500, year=2025), _kpi(1, 12000, 11000, year=2026)]
    label, _ = compute_prop_projection(kpis)
    assert label == "2026"


def test_projection_q1_plus_q2q4_budget():
    kpis = [
        _kpi(1, 10000, 9000), _kpi(2, 10000, 9000), _kpi(3, 10000, 9000),
        _kpi(4, None, 9000),  _kpi(5, None, 9000),  _kpi(6, None, 9000),
        _kpi(7, None, 9000),  _kpi(8, None, 9000),  _kpi(9, None, 9000),
        _kpi(10, None, 9000), _kpi(11, None, 9000), _kpi(12, None, 9000),
    ]
    label, proj = compute_prop_projection(kpis)
    ai = proj["actual_income"]
    # Q1 actual income = 30000; Q2-Q4 budget income = 9 × 9000 = 81000; total = 111000
    assert ai["proj_fy"] == pytest.approx(111000)
    assert ai["fy_budget"] == pytest.approx(108000)  # 12 × 9000


def test_fallback_when_no_q2q4_budget():
    # Only Q1 data present — falls back to Q1 budget × 3 for Q2-Q4
    kpis = [_kpi(1, 10000, 9000), _kpi(2, 10000, 9000), _kpi(3, 10000, 9000)]
    _, proj = compute_prop_projection(kpis)
    ai = proj["actual_income"]
    # Q1 actual = 30000; Q2-Q4 fallback = 9000 × 3 = 27000; proj_fy = 57000
    assert ai["proj_fy"] == pytest.approx(57000)
    # FY budget fallback = 9000 × 4 = 36000
    assert ai["fy_budget"] == pytest.approx(36000)
```

- [ ] **Step 2: Run tests to confirm they fail**

```powershell
pytest tests/test_projection.py -v
```

Expected: All FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Create `app/ui/projection.py`**

```python
"""Full-year projection helper — extracted from routes/property_detail.py."""


def compute_prop_projection(prop_kpis: list[dict]) -> tuple[str, dict]:
    """
    Compute full-year projection for a single property's KPI list.
    Returns (proj_yr_label, projection_dict).

    projection_dict keys: "actual_income", "actual_expenses", "actual_noi"
    Each value is a dict: {q1_actual, proj_fy, fy_budget, var_to_plan, var_to_plan_pct}
    """
    if not prop_kpis:
        return "", {}

    years_present = {k["year"] for k in prop_kpis if k.get("year")}
    if not years_present:
        return "", {}

    proj_yr = max(years_present)
    proj_yr_label = str(proj_yr)

    def _psum(kpi_list, field):
        vals = [k.get(field) for k in kpi_list if k.get(field) is not None]
        return sum(vals) if vals else None

    q1k   = [k for k in prop_kpis if k.get("year") == proj_yr and k.get("month") in (1, 2, 3)]
    q2q4k = [k for k in prop_kpis if k.get("year") == proj_yr and k.get("month") in range(4, 13)]
    ayk   = [k for k in prop_kpis if k.get("year") == proj_yr]

    projection = {}
    for pk, bk in [("actual_income",   "budget_income"),
                   ("actual_expenses",  "budget_expenses"),
                   ("actual_noi",       "budget_noi")]:
        q1_act   = _psum(q1k,   pk)
        q2q4_bud = _psum(q2q4k, bk)
        ay_bud   = _psum(ayk,   bk)

        if not q2q4_bud:
            q1_bud   = _psum(q1k, bk)
            q2q4_bud = (q1_bud * 3) if q1_bud is not None else None
            fy_bud   = (q1_bud * 4) if q1_bud is not None else None
        else:
            fy_bud = ay_bud

        proj_fy = (q1_act + q2q4_bud) if (q1_act is not None and q2q4_bud is not None) else None
        var     = (proj_fy - fy_bud)   if (proj_fy is not None and fy_bud is not None) else None
        var_pct = (var / abs(fy_bud))  if (var is not None and fy_bud) else None

        projection[pk] = {
            "q1_actual":       q1_act,
            "proj_fy":         proj_fy,
            "fy_budget":       fy_bud,
            "var_to_plan":     var,
            "var_to_plan_pct": var_pct,
        }

    return proj_yr_label, projection
```

- [ ] **Step 4: Run tests to confirm they pass**

```powershell
pytest tests/test_projection.py -v
```

Expected: All 5 PASS.

- [ ] **Step 5: Commit**

```powershell
git add app/ui/projection.py tests/test_projection.py
git commit -m "feat: extract projection helper to app/ui/projection"
```

---

## Task 6: Extract Pipeline Function

**Files:**
- Create: `app/ui/pipeline.py`
- Create: `tests/test_pipeline.py`

This is the largest extraction. `run_analysis_pipeline()` contains all the business logic from `app/routes/upload.py` with Flask-specific code (`request.files`, `secure_filename`, `redirect`) replaced by plain Python parameters.

- [ ] **Step 1: Write failing tests**

Create `tests/test_pipeline.py`:

```python
"""Integration tests for the analysis pipeline — replaces test_routes.py."""
import io
import os
import pytest
import openpyxl

from app.ui.pipeline import run_analysis_pipeline


def _make_fin_workbook_bytes(sheet_title="Actual - Test Property"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(["Account", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    ws.append(["Gross Potential Rent",
               5000, 5000, 5000, 5000, 5000, 5000,
               5000, 5000, 5000, 5000, 5000, 5000])
    ws.append(["Vacancy Loss",
               -250, -250, -250, -250, -250, -250,
               -250, -250, -250, -250, -250, -250])
    ws.append(["Management Fee",
               500, 500, 500, 500, 500, 500,
               500, 500, 500, 500, 500, 500])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


@pytest.fixture(autouse=True)
def chdir_tmp(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("runs", exist_ok=True)


def test_pipeline_returns_run_id():
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[],
        ar_files=[],
        settings={"portfolio_name": "Test Portfolio", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": ["PM One"],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    assert run_id
    assert os.path.isdir(os.path.join("runs", run_id))


def test_pipeline_writes_metadata():
    import json
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "My Portfolio", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    meta_path = os.path.join("runs", run_id, "metadata.json")
    assert os.path.isfile(meta_path)
    with open(meta_path) as f:
        meta = json.load(f)
    assert meta["portfolio_name"] == "My Portfolio"


def test_pipeline_writes_main_workbook():
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "WB Test", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(), "carveout_properties": set(),
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    run_dir = os.path.join("runs", run_id)
    xlsx_files = [f for f in os.listdir(run_dir) if f.endswith(".xlsx")]
    assert len(xlsx_files) == 2


def test_pipeline_no_files_raises():
    with pytest.raises(ValueError, match="No valid"):
        run_analysis_pipeline(
            fin_files=[], occ_files=[], ar_files=[],
            settings={"portfolio_name": "Empty", "eco_occ_target": 0.95,
                      "use_budget_eco_occ": False, "pm_names": [],
                      "excluded_properties": set(), "carveout_properties": set(),
                      "stabilized_properties": set(), "period_filter": "Full Year",
                      "selected_months": [], "custom_mapping": None},
        )


def test_pipeline_carveout_property_flagged():
    from app.storage.runs import load_run
    run_id = run_analysis_pipeline(
        fin_files=[("test.xlsx", _make_fin_workbook_bytes())],
        occ_files=[], ar_files=[],
        settings={"portfolio_name": "CO Test", "eco_occ_target": 0.95,
                  "use_budget_eco_occ": False, "pm_names": [],
                  "excluded_properties": set(),
                  "carveout_properties": {"test property"},
                  "stabilized_properties": set(), "period_filter": "Full Year",
                  "selected_months": [], "custom_mapping": None},
    )
    data = load_run(run_id)
    carveout_kpis = [k for k in data["kpis"] if k.get("is_carveout")]
    assert len(carveout_kpis) > 0
```

- [ ] **Step 2: Run tests to confirm they fail**

```powershell
pytest tests/test_pipeline.py -v
```

Expected: All FAIL with `ModuleNotFoundError`.

- [ ] **Step 3: Create `app/ui/pipeline.py`**

```python
"""Analysis pipeline — Flask-free version of app/routes/upload.py business logic.

Accepts file bytes instead of Flask request.files objects. Returns run_id.
"""
import io
import csv
import os
import re
import tempfile
from datetime import datetime

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
from config import (QUARTERS, PROPERTY_NAME_MAP, MONTHS,
                    PERMANENT_EXCLUSIONS, PROPERTY_METADATA, PROPERTY_PM_EXCLUSIONS)


def _safe_filename(name: str) -> str:
    """Sanitize a filename (replaces werkzeug.utils.secure_filename)."""
    name = os.path.basename(name)
    name = re.sub(r"[^\w\s\-.]", "", name).strip()
    return name or "upload"


def _save_bytes_to_temp(filename: str, data: bytes) -> str:
    """Write bytes to a NamedTemporaryFile and return the path."""
    suffix = os.path.splitext(filename)[1] or ".xlsx"
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path


def _detect_partial_year(kpis) -> set[str]:
    """Auto-detect properties with fewer months than the max in their year."""
    from collections import defaultdict
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
        if i < len(pm_names):
            pm_name_map[os.path.basename(path)] = pm_names[i]

    if not saved_paths:
        raise ValueError("No valid .xlsx files were provided.")

    # ── Save occupancy and AR files ───────────────────────────────────────────
    occ_paths = [_save_bytes_to_temp(fn, d) for fn, d in occ_files if fn]
    ar_paths  = [_save_bytes_to_temp(fn, d) for fn, d in ar_files  if fn]

    try:
        # ── Parse ──────────────────────────────────────────────────────────────
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
            rows = parse_ar_aging_reports(ar_paths)
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
```

- [ ] **Step 4: Run tests to confirm they pass**

```powershell
pytest tests/test_pipeline.py -v
```

Expected: All 5 PASS.

- [ ] **Step 5: Run the full test suite to confirm nothing regressed**

```powershell
pytest tests/ -q --ignore=tests/test_routes.py
```

Expected: All tests pass.

- [ ] **Step 6: Commit**

```powershell
git add app/ui/pipeline.py tests/test_pipeline.py
git commit -m "feat: extract analysis pipeline to app/ui/pipeline (Flask-free)"
```

---

## Task 7: Upload Page

**Files:**
- Modify: `pages/upload.py` (replace placeholder)

- [ ] **Step 1: Replace `pages/upload.py` with the full implementation**

```python
"""Upload page — collects files and settings, runs the analysis pipeline."""
import io
import csv
import streamlit as st
from config import ECO_OCC_TARGET, QUARTERS, MONTHS

st.header("New Portfolio Analysis")
st.caption("Upload financial statement workbooks to generate a KPI analysis.")

# ── File uploads ──────────────────────────────────────────────────────────────
with st.expander("📁 Required Files", expanded=True):
    fin_files = st.file_uploader(
        "Financial Statement Workbooks (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="One or more 12-month income statement workbooks. "
             "Actual, Budget, or Actual vs. Budget formats accepted.",
    )
    portfolio_name = st.text_input(
        "Portfolio Name", value="Portfolio",
        help="Used as the workbook title and download filename.",
    )

with st.expander("📋 Optional Files", expanded=False):
    occ_files = st.file_uploader(
        "Physical Occupancy Report (.xlsx) — optional",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Required for Physical Occ %, Leakage Gap, and Per Unit calculations. "
             "Columns: Property, Year, Month, Occupied Units, Total Units.",
    )
    ar_files = st.file_uploader(
        "AR Aging Reports (.xlsx) — optional",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
    mapping_file = st.file_uploader(
        "Custom Account Mapping (.csv) — optional",
        type=["csv"],
        help="CSV columns: account_name, assigned_category, treatment, "
             "include_in_noi, include_in_eco_occ",
    )

# ── Analysis settings ─────────────────────────────────────────────────────────
with st.expander("⚙️ Analysis Settings", expanded=False):
    eco_occ_target_pct = st.number_input(
        "Economic Occupancy Target (%)", min_value=0.0, max_value=100.0,
        value=float(ECO_OCC_TARGET * 100), step=0.5,
    )
    use_budget_eco_occ = st.checkbox(
        "Use Budget Eco Occ % as target (instead of fixed %)",
        value=False,
    )
    pm_names_raw = st.text_area(
        "Property Manager Names (one per line, matches file order)",
        height=80,
        help="If blank, PM name is inferred from the filename.",
    )
    period_filter = st.selectbox(
        "Reporting Period", ["Full Year", "Q1", "Q2", "Q3", "Q4", "Selected Months"],
    )
    selected_months = []
    if period_filter == "Selected Months":
        month_names = [MONTHS[i] for i in range(1, 13)]
        selected_month_names = st.multiselect("Select Months", month_names)
        selected_months = [i for i in range(1, 13) if MONTHS[i] in selected_month_names]

with st.expander("🔧 Advanced Settings", expanded=False):
    excluded_raw   = st.text_area("Excluded Properties (one per line)", height=80)
    carveout_raw   = st.text_area(
        "Carve-out Properties (one per line)",
        height=80,
        help="Shown in property detail but excluded from portfolio totals.",
    )
    stabilized_raw = st.text_area(
        "Recently Stabilised Properties (one per line)",
        height=80,
        help="Excluded from portfolio YoY comparisons.",
    )

# ── Run button ────────────────────────────────────────────────────────────────
if st.button("🚀 Run Analysis", type="primary", use_container_width=True):
    if not fin_files:
        st.error("Please upload at least one financial statement workbook.")
        st.stop()

    # Parse custom mapping CSV if uploaded
    custom_mapping = None
    if mapping_file:
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

    settings = {
        "portfolio_name":       portfolio_name.strip() or "Portfolio",
        "eco_occ_target":       eco_occ_target_pct / 100.0,
        "use_budget_eco_occ":   use_budget_eco_occ,
        "pm_names":             [l.strip() for l in pm_names_raw.splitlines() if l.strip()],
        "excluded_properties":  {p.strip().lower() for p in excluded_raw.splitlines() if p.strip()},
        "carveout_properties":  {p.strip().lower() for p in carveout_raw.splitlines() if p.strip()},
        "stabilized_properties":{p.strip() for p in stabilized_raw.splitlines() if p.strip()},
        "period_filter":        period_filter,
        "selected_months":      selected_months,
        "custom_mapping":       custom_mapping,
    }

    from app.ui.pipeline import run_analysis_pipeline

    with st.spinner("Running analysis — this may take 30–60 seconds…"):
        try:
            fin_bytes = [(f.name, f.read()) for f in fin_files]
            occ_bytes = [(f.name, f.read()) for f in (occ_files or [])]
            ar_bytes  = [(f.name, f.read()) for f in (ar_files  or [])]
            run_id = run_analysis_pipeline(fin_bytes, occ_bytes, ar_bytes, settings)
        except ValueError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"Analysis failed: {e}")
            st.stop()

    st.session_state["current_run_id"] = run_id
    st.success(f"Analysis complete! Run ID: `{run_id}`")
    st.switch_page("pages/dashboard.py")
```

- [ ] **Step 2: Start the app and manually verify the upload page**

```powershell
streamlit run streamlit_app.py
```

Verify:
- Upload page loads with all expanders
- "Required Files" expander is open by default
- Clicking "Run Analysis" with no files shows the error message
- No Python errors in the terminal

Stop the server.

- [ ] **Step 3: Commit**

```powershell
git add pages/upload.py
git commit -m "feat: Streamlit upload page with pipeline wiring"
```

---

## Task 8: Dashboard Page

**Files:**
- Modify: `pages/dashboard.py` (replace placeholder)

- [ ] **Step 1: Replace `pages/dashboard.py` with the full implementation**

```python
"""Dashboard page — portfolio KPI summary, property table, AR aging, NOI rankings."""
import io
import zipfile
import streamlit as st
import pandas as pd

from app.storage.runs import load_run, list_runs
from app.ui.formatting import fmt_currency, fmt_pct
from app.ui.aggregation import (agg_kpis, agg_ar, agg_ar_for_prop,
                                  ar_yoy_delta, pct_delta,
                                  month_to_quarter, quarter_label, quarter_months,
                                  ar_period_label)
from app.ui.kpi_definitions import (SUMMARY_KPI_DEFINITIONS, KPI_TOOLTIPS,
                                     BUDGET_YOY_KEY, YOY_CURRENCY_KEYS,
                                     YOY_FAVORABLE_IF_POSITIVE, PCT_VARIANCE_THRESHOLD_KEYS)
from config import ECO_OCC_TARGET


# ── Load run ──────────────────────────────────────────────────────────────────

@st.cache_data
def _load(run_id: str) -> dict:
    return load_run(run_id)


run_id = st.session_state.get("current_run_id")

if not run_id:
    # Allow selecting a run from history if none is active
    runs = list_runs()
    if not runs:
        st.info("No analysis loaded. Go to **New Analysis** to upload files.")
        st.stop()
    options = {f"{r['portfolio_name']} ({r['created_at'][:10]})": r["run_id"] for r in runs}
    choice = st.selectbox("Select a previous run:", list(options.keys()))
    run_id = options[choice]
    st.session_state["current_run_id"] = run_id

data = _load(run_id)
meta = data["metadata"]
kpis = data["kpis"]

portfolio_name      = meta.get("portfolio_name", "Portfolio")
eco_occ_target      = meta.get("eco_occ_target", ECO_OCC_TARGET)
use_budget_eco_occ  = meta.get("use_budget_eco_occ", False)
years               = meta.get("years", [])
props               = meta.get("properties", [])
partial_year_props  = set(meta.get("partial_year_properties", []))
num_props           = meta.get("num_properties", len(props))

# ── Header ────────────────────────────────────────────────────────────────────
col_title, col_dl = st.columns([4, 1])
with col_title:
    st.title(f"{portfolio_name}")
    st.caption(f"{num_props} Properties  ·  {', '.join(str(y) for y in years)}")
with col_dl:
    # Build download ZIP in memory
    run_dir   = __import__("os").path.join("runs", run_id)
    main_wb   = __import__("os").path.join(run_dir, meta.get("main_workbook", ""))
    backup_wb = __import__("os").path.join(run_dir, meta.get("backup_workbook", ""))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if __import__("os").path.isfile(main_wb):
            zf.write(main_wb, __import__("os").path.basename(main_wb))
        if __import__("os").path.isfile(backup_wb):
            zf.write(backup_wb, __import__("os").path.basename(backup_wb))
    buf.seek(0)
    st.download_button(
        "⬇ Download Workbooks (.zip)",
        data=buf,
        file_name=f"{portfolio_name} Analysis Workbooks.zip",
        mime="application/zip",
        use_container_width=True,
    )

# ── Quarter aggregation ───────────────────────────────────────────────────────
all_quarters: set[tuple] = set()
for k in kpis:
    if not k.get("is_carveout") and k.get("year") and k.get("month"):
        all_quarters.add((k["year"], month_to_quarter(k["month"])))
sorted_quarters = sorted(all_quarters, reverse=True)
period_labels = [quarter_label(yr, q) for (yr, q) in sorted_quarters]

period_aggs: dict[str, dict] = {}
for (yr, q) in sorted_quarters:
    months = quarter_months(q)
    q_kpis = [k for k in kpis
              if k.get("year") == yr and k.get("month") in months
              and not k.get("is_carveout") and not k.get("is_partial_year")]
    lbl = quarter_label(yr, q)
    period_aggs[lbl] = agg_kpis(q_kpis)

latest_period_label = period_labels[0] if period_labels else ""

years_sorted = sorted({k["year"] for k in kpis
                        if not k.get("is_carveout") and k.get("year")})
year_aggs: dict[int, dict] = {}
for yr in years_sorted:
    yr_kpis = [k for k in kpis if k.get("year") == yr
               and not k.get("is_carveout") and not k.get("is_partial_year")]
    year_aggs[yr] = agg_kpis(yr_kpis)

year_pairs = list(reversed([(years_sorted[i], years_sorted[i + 1])
                              for i in range(len(years_sorted) - 1)]))

# ── Full-year projection ──────────────────────────────────────────────────────
proj_yr       = max(years_sorted) if years_sorted else None
proj_yr_label = str(proj_yr) if proj_yr else ""
projection_data: dict[str, dict] = {}

if proj_yr:
    _q1k   = [k for k in kpis if k.get("year") == proj_yr and k.get("month") in {1,2,3}
              and not k.get("is_carveout") and not k.get("is_partial_year")]
    _q2q4k = [k for k in kpis if k.get("year") == proj_yr and k.get("month") in range(4,13)
              and not k.get("is_carveout") and not k.get("is_partial_year")]
    _ayk   = [k for k in kpis if k.get("year") == proj_yr
              and not k.get("is_carveout") and not k.get("is_partial_year")]
    q1a    = agg_kpis(_q1k)
    q2q4a  = agg_kpis(_q2q4k)
    aya    = agg_kpis(_ayk)
    for pk, bk in [("actual_income","budget_income"),
                   ("actual_expenses","budget_expenses"),
                   ("actual_noi","budget_noi")]:
        q1_act   = q1a.get(pk)
        q2q4_bud = q2q4a.get(bk)
        fy_bud   = aya.get(bk) if q2q4_bud else (q1a.get(bk) * 4 if q1a.get(bk) else None)
        if not q2q4_bud:
            q1_bud   = q1a.get(bk)
            q2q4_bud = (q1_bud * 3) if q1_bud is not None else None
            fy_bud   = (q1_bud * 4) if q1_bud is not None else None
        proj_fy = (q1_act + q2q4_bud) if (q1_act is not None and q2q4_bud is not None) else None
        var     = (proj_fy - fy_bud)   if (proj_fy is not None and fy_bud is not None) else None
        projection_data[pk] = {"proj_fy": proj_fy, "fy_budget": fy_bud, "var_to_plan": var}


# ── Helper: build styled DataFrame for KPI table ─────────────────────────────
def _build_kpi_df(period_labels, period_aggs):
    """Build a formatted string DataFrame from SUMMARY_KPI_DEFINITIONS."""
    rows = []
    for defn in SUMMARY_KPI_DEFINITIONS:
        if defn is None:
            rows.append({"KPI": "", **{lbl: "" for lbl in period_labels}})
            continue
        label, key, fmt, fav, group_id, is_hdr = defn
        row = {"KPI": ("▶ " if is_hdr else "   ") + label}
        for lbl in period_labels:
            val = period_aggs.get(lbl, {}).get(key)
            row[lbl] = fmt_currency(val) if fmt == "currency" else fmt_pct(val)
        rows.append(row)
    return pd.DataFrame(rows).set_index("KPI")


def _color_variance_col(col_series, favorable_positive):
    """Return per-cell CSS styles for a variance column."""
    styles = []
    for label, val_str in col_series.items():
        if not label.strip() or val_str in ("—", ""):
            styles.append("")
            continue
        try:
            # Strip formatting to get numeric value
            numeric = float(val_str.replace("$", "").replace(",", "")
                            .replace("(", "-").replace(")", "")
                            .replace("%", ""))
        except (ValueError, AttributeError):
            styles.append("")
            continue
        if favorable_positive:
            styles.append("color: #059669; font-weight:600" if numeric > 0
                          else "color: #dc2626; font-weight:600" if numeric < 0 else "")
        else:
            styles.append("color: #059669; font-weight:600" if numeric < 0
                          else "color: #dc2626; font-weight:600" if numeric > 0 else "")
    return styles


# ── Portfolio Summary KPI table ───────────────────────────────────────────────
st.subheader(f"Portfolio Summary — {num_props} Properties")

if period_labels:
    kpi_df = _build_kpi_df(period_labels, period_aggs)

    # Apply variance coloring to variance columns
    variance_cols = {
        lbl: defn for defn in SUMMARY_KPI_DEFINITIONS if defn is not None
        for lbl in [defn[0]]
        if defn[3] is not None and defn[0] in kpi_df.columns  # has favorable_positive
    }
    styler = kpi_df.style
    for defn in SUMMARY_KPI_DEFINITIONS:
        if defn is None:
            continue
        label, key, fmt, fav, group_id, is_hdr = defn
        if fav is not None and label in kpi_df.columns:
            # Color the variance column across all periods
            for lbl in period_labels:
                if lbl in kpi_df.columns:
                    pass  # applied row-by-row below

    # Bold group header rows
    def _style_row(row):
        label_clean = row.name.strip()
        if label_clean.startswith("▶"):
            return ["font-weight: bold; background-color: #DEEAF1"] * len(row)
        if not label_clean:
            return ["background-color: #F4F6F9"] * len(row)
        return [""] * len(row)

    styled = kpi_df.style.apply(_style_row, axis=1)
    st.dataframe(styled, use_container_width=True, height=600)
else:
    st.info("No period data available.")

# ── Full-year projection summary ──────────────────────────────────────────────
if proj_yr and projection_data:
    with st.expander(f"📈 Full-Year {proj_yr_label} Projection (Q1 Actual + Q2–Q4 Budget)", expanded=False):
        proj_rows = []
        for label, pk in [("Income", "actual_income"),
                           ("Expenses", "actual_expenses"),
                           ("NOI", "actual_noi")]:
            pd_row = projection_data.get(pk, {})
            proj_rows.append({
                "Metric":             label,
                "Projected Full Year": fmt_currency(pd_row.get("proj_fy")),
                "FY Budget":           fmt_currency(pd_row.get("fy_budget")),
                "Variance to Plan":    fmt_currency(pd_row.get("var_to_plan")),
            })
        st.dataframe(pd.DataFrame(proj_rows).set_index("Metric"), use_container_width=True)

# ── NOI vs Budget ranking ─────────────────────────────────────────────────────
if sorted_quarters:
    _lq_yr, _lq = sorted_quarters[0]
    lq_label = quarter_label(_lq_yr, _lq)
    _lq_months = quarter_months(_lq)
    _vb_rows = []
    for prop in sorted(props):
        pq = [k for k in kpis
              if k.get("property_name") == prop
              and k.get("year") == _lq_yr and k.get("month") in _lq_months
              and not k.get("is_carveout") and not k.get("is_partial_year")]
        if not pq:
            continue
        ag = agg_kpis(pq)
        an = ag.get("actual_noi")
        bn = ag.get("budget_noi")
        if an is None or bn is None:
            continue
        nv = an - bn
        nv_pct = nv / abs(bn) if bn else None
        _vb_rows.append({
            "Property": prop,
            "Actual NOI": fmt_currency(an),
            "Budget NOI": fmt_currency(bn),
            "NOI Variance": fmt_currency(nv),
            "Var %": fmt_pct(nv_pct),
            "_nv": nv,
        })
    _vb_rows.sort(key=lambda r: r["_nv"])

    st.subheader(f"NOI vs Budget — {lq_label}")
    col_top, col_bot = st.columns(2)
    with col_top:
        st.markdown("**✅ Top 5 Above Budget**")
        top5 = [r for r in reversed(_vb_rows)][:5]
        if top5:
            df_top = pd.DataFrame(top5).drop(columns=["_nv"]).set_index("Property")
            st.dataframe(df_top, use_container_width=True)
        else:
            st.caption("No data")
    with col_bot:
        st.markdown("**⚠️ Top 5 Below Budget**")
        bot5 = _vb_rows[:5]
        if bot5:
            df_bot = pd.DataFrame(bot5).drop(columns=["_nv"]).set_index("Property")
            st.dataframe(df_bot, use_container_width=True)
        else:
            st.caption("No data")

# ── Property table ────────────────────────────────────────────────────────────
st.subheader(f"Properties — {num_props} in Analysis")

latest_yr = max(years) if years else None
prop_rows = []
partial_yr_rows = []

for prop in sorted(props):
    is_py = prop in partial_year_props
    prop_kpis = [k for k in kpis
                 if k.get("property_name") == prop and k.get("year") == latest_yr]
    if not prop_kpis:
        continue
    ag = agg_kpis(prop_kpis)
    row = {
        "Property":    prop,
        "PM":          prop_kpis[0].get("pm_name", ""),
        "Actual NOI":  fmt_currency(ag.get("actual_noi")),
        "NOI Var":     fmt_currency(ag.get("noi_variance")),
        "Eco Occ %":   fmt_pct(ag.get("eco_occ_pct")),
        "Phys Occ %":  fmt_pct(ag.get("physical_occ_pct")),
        "Leakage":     fmt_pct(ag.get("leakage_gap")),
        "NOI/Unit":    fmt_currency(ag.get("noi_per_unit")),
        "_prop":       prop,
    }
    if is_py:
        partial_yr_rows.append(row)
    else:
        prop_rows.append(row)

def _render_prop_table(rows, label):
    if not rows:
        return
    st.markdown(f"**{label}**")
    df = pd.DataFrame(rows).drop(columns=["_prop"])
    df = df.set_index("Property")
    # Add a View button column — Streamlit can't put buttons in dataframes,
    # so we show the table and then a selectbox for navigation.
    st.dataframe(df, use_container_width=True)

_render_prop_table(prop_rows, f"Full-Year Properties ({len(prop_rows)})")

# Property detail navigation
all_prop_names = [r["_prop"] for r in prop_rows + partial_yr_rows]
if all_prop_names:
    selected = st.selectbox("View property detail:", ["— select —"] + all_prop_names)
    if selected != "— select —":
        st.session_state["selected_property"] = selected
        st.switch_page("pages/property_detail.py")

if partial_yr_rows:
    with st.expander(f"Recently Stabilised / Partial-Year Properties ({len(partial_yr_rows)})", expanded=False):
        _render_prop_table(partial_yr_rows, "")

# ── AR Aging summary ──────────────────────────────────────────────────────────
ar_rows = data.get("ar_aging", [])
if ar_rows:
    st.subheader("AR Aging Summary")

    _bd: dict[tuple, float] = {}
    for _k in kpis:
        key = (_k["property_name"], _k["year"], _k["month"])
        _bd[key] = (_bd.get(key) or 0.0) + (_k.get("bad_debt") or 0.0)

    for rtype in ["Tenant Rent", "Subsidy"]:
        periods = sorted({(r["year"], r["month"]) for r in ar_rows
                          if r["receivable_type"] == rtype}, reverse=True)
        if not periods:
            continue
        with st.expander(f"{rtype} AR", expanded=True):
            ar_summary_rows = []
            for (yr, mo) in periods[:6]:  # Show latest 6 periods
                ag = agg_ar(ar_rows, rtype, yr, mo)
                if not ag:
                    continue
                period_props = {r["property_name"] for r in ar_rows
                                if r["receivable_type"] == rtype
                                and r["year"] == yr and r["month"] == mo}
                bd_period = sum(_bd.get((p, yr, mo), 0.0) for p in period_props) or None
                ar_summary_rows.append({
                    "Period":         ar_period_label(yr, mo),
                    "# Props":        ag["property_count"],
                    "Current Owed":   fmt_currency(ag["current_owed"]),
                    "Pre-payments":   fmt_currency(ag["prepayments"]),
                    "% >60 Days":     fmt_pct(ag["pct_overdue"]),
                    "Bad Debt (W/O)": fmt_currency(bd_period),
                })
            if ar_summary_rows:
                st.dataframe(
                    pd.DataFrame(ar_summary_rows).set_index("Period"),
                    use_container_width=True,
                )

# ── Quality checks ────────────────────────────────────────────────────────────
quality_checks = data.get("quality_checks", [])
if quality_checks:
    with st.expander("🔍 Quality Checks", expanded=False):
        qc_rows = [{"Check": qc["check_name"],
                    "Status": "✅ PASS" if qc["passed"] else "❌ FAIL",
                    "Detail": qc["detail"]} for qc in quality_checks]
        st.dataframe(pd.DataFrame(qc_rows).set_index("Check"), use_container_width=True)
```

- [ ] **Step 2: Manually test the dashboard loads after a run**

```powershell
streamlit run streamlit_app.py
```

- Upload a real financial workbook on the New Analysis page
- Click "Run Analysis"
- Verify it navigates to the dashboard
- Verify the KPI table renders with period columns
- Verify the property table appears
- Verify the download button appears and clicking it downloads a ZIP

Stop the server.

- [ ] **Step 3: Commit**

```powershell
git add pages/dashboard.py
git commit -m "feat: Streamlit dashboard page (KPI table, property table, AR, rankings)"
```

---

## Task 9: Property Detail Page

**Files:**
- Modify: `pages/property_detail.py` (replace placeholder)

- [ ] **Step 1: Replace `pages/property_detail.py` with the full implementation**

```python
"""Property detail page — monthly KPIs, full-year projection, AR aging."""
import streamlit as st
import pandas as pd

from app.storage.runs import load_run
from app.ui.formatting import fmt_currency, fmt_pct
from app.ui.projection import compute_prop_projection


@st.cache_data
def _load(run_id: str) -> dict:
    return load_run(run_id)


run_id          = st.session_state.get("current_run_id")
property_name   = st.session_state.get("selected_property")

if not run_id or not property_name:
    st.info("No property selected. Return to the **Dashboard** and choose a property.")
    if st.button("← Back to Dashboard"):
        st.switch_page("pages/dashboard.py")
    st.stop()

data = _load(run_id)

# ── Header ────────────────────────────────────────────────────────────────────
col_back, col_title = st.columns([1, 6])
with col_back:
    if st.button("← Dashboard"):
        st.switch_page("pages/dashboard.py")
with col_title:
    st.title(property_name)

# ── Monthly KPI table ─────────────────────────────────────────────────────────
prop_kpis = [k for k in data["kpis"] if k["property_name"] == property_name]
if not prop_kpis:
    st.error("No KPI data found for this property.")
    st.stop()

prop_kpis.sort(key=lambda k: (-k["year"], -k["month"]))

_MONTH_ABBR = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
               7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}

st.subheader("Monthly KPIs")
kpi_rows = []
for k in prop_kpis:
    period_str = f"{k['year']} {_MONTH_ABBR.get(k.get('month', 0), str(k.get('month', '?')))}"
    noi_var = k.get("noi_variance")
    kpi_rows.append({
        "Period":       period_str,
        "Income":       fmt_currency(k.get("actual_income")),
        "Expenses":     fmt_currency(k.get("actual_expenses")),
        "NOI":          fmt_currency(k.get("actual_noi")),
        "NOI Var":      fmt_currency(noi_var),
        "Eco Occ %":    fmt_pct(k.get("eco_occ_pct")),
        "Phys Occ %":   fmt_pct(k.get("physical_occ_pct")),
        "Leakage":      fmt_pct(k.get("leakage_gap")),
        "Income/Unit":  fmt_currency(k.get("income_per_unit")),
        "Exp/Unit":     fmt_currency(k.get("expense_per_unit")),
        "NOI/Unit":     fmt_currency(k.get("noi_per_unit")),
        "GPR":          fmt_currency(k.get("gpr")),
        "Vacancy":      fmt_currency(k.get("vacancy")),
        "Concessions":  fmt_currency(k.get("concessions")),
        "Bad Debt":     fmt_currency(k.get("bad_debt")),
        "_noi_var":     noi_var,
    })

df_kpi = pd.DataFrame(kpi_rows).set_index("Period")

def _color_noi_var(col):
    styles = []
    for v in col:
        try:
            num = float(str(v).replace("$", "").replace(",", "")
                        .replace("(", "-").replace(")", ""))
            styles.append("color: #059669; font-weight:600" if num > 0
                          else "color: #dc2626; font-weight:600" if num < 0 else "")
        except (ValueError, AttributeError):
            styles.append("")
    return styles

df_display = df_kpi.drop(columns=["_noi_var"])
styled = df_display.style.apply(_color_noi_var, subset=["NOI Var"])
st.dataframe(styled, use_container_width=True)

# ── Full-year projection ──────────────────────────────────────────────────────
proj_yr_label, prop_projection = compute_prop_projection(prop_kpis)

if prop_projection and proj_yr_label:
    st.subheader(f"Full Year {proj_yr_label} Projection")
    st.caption(
        f"Projected Full Year = Q1 {proj_yr_label} Actual + Q2–Q4 {proj_yr_label} Budget "
        f"(fallback: Q1 Budget × 3 if Q2–Q4 budget not available)."
    )
    proj_rows = []
    for label, pk in [("Income", "actual_income"),
                       ("Expenses", "actual_expenses"),
                       ("NOI", "actual_noi")]:
        pd_row = prop_projection.get(pk, {})
        var     = pd_row.get("var_to_plan")
        var_pct = pd_row.get("var_to_plan_pct")
        proj_rows.append({
            "Metric":              label,
            "Q1 Actual":           fmt_currency(pd_row.get("q1_actual")),
            "Projected Full Year": fmt_currency(pd_row.get("proj_fy")),
            "FY Budget":           fmt_currency(pd_row.get("fy_budget")),
            "Variance to Plan":    fmt_currency(var),
            "Var %":               fmt_pct(var_pct),
        })
    st.dataframe(pd.DataFrame(proj_rows).set_index("Metric"), use_container_width=True)

# ── AR Aging detail ───────────────────────────────────────────────────────────
raw_ar = data.get("ar_aging", [])
prop_ar = [r for r in raw_ar if r["property_name"] == property_name]

if prop_ar:
    _MONTH_ABBR2 = _MONTH_ABBR

    for r in prop_ar:
        r["total_overdue"] = r["owed_31_60"] + r["owed_61_90"] + r["owed_over_90"]
        charge = r.get("charge_amount", 0)
        r["pct_overdue"] = (r["total_overdue"] / charge) if charge and charge > 0 else None
        r["period_label"] = f"{_MONTH_ABBR2.get(r['month'], str(r['month']))}-{r['year']}"

    prop_ar.sort(key=lambda r: (r["receivable_type"], -r["year"], -r["month"]))

    ar_by_type: dict = {}
    for r in prop_ar:
        ar_by_type.setdefault(r["receivable_type"], []).append(r)

    st.subheader("AR Aging Detail")
    for rtype, rows in ar_by_type.items():
        with st.expander(rtype, expanded=True):
            ar_table = []
            for r in rows:
                ar_table.append({
                    "Period":       r["period_label"],
                    "Charge Amt":   fmt_currency(r.get("charge_amount")),
                    "Current":      fmt_currency(r.get("current_owed")),
                    "0–30":         fmt_currency(r.get("owed_0_30")),
                    "31–60":        fmt_currency(r.get("owed_31_60")),
                    "61–90":        fmt_currency(r.get("owed_61_90")),
                    "Over 90":      fmt_currency(r.get("owed_over_90")),
                    "Pre-payments": fmt_currency(r.get("prepayments")),
                    "% >30 Days":   fmt_pct(r.get("pct_overdue")),
                })
            st.dataframe(
                pd.DataFrame(ar_table).set_index("Period"),
                use_container_width=True,
            )
```

- [ ] **Step 2: Manually test property detail navigation**

```powershell
streamlit run streamlit_app.py
```

- Upload a workbook and run the analysis
- On the dashboard, select a property from the dropdown
- Verify the property detail page loads with the monthly KPI table
- Verify "← Dashboard" button returns to the dashboard

Stop the server.

- [ ] **Step 3: Commit**

```powershell
git add pages/property_detail.py
git commit -m "feat: Streamlit property detail page (KPIs, projection, AR aging)"
```

---

## Task 10: History Page

**Files:**
- Modify: `pages/history.py` (replace placeholder)

- [ ] **Step 1: Replace `pages/history.py` with the full implementation**

```python
"""History page — list of past runs with view, download, and delete actions."""
import io
import os
import zipfile
import streamlit as st
import pandas as pd

from app.storage.runs import list_runs, delete_run, load_run

st.header("Analysis History")
st.caption("Click **View** to reload a past analysis. Click **Delete** to permanently remove it.")

runs = list_runs()

if not runs:
    st.info("No analyses yet. Go to **New Analysis** to upload files.")
    st.stop()

for run in runs:
    run_id   = run["run_id"]
    name     = run.get("portfolio_name", "—")
    created  = run.get("created_at", "")[:10]
    num_p    = run.get("num_properties", "—")
    yrs      = ", ".join(str(y) for y in run.get("years", []))
    pms      = ", ".join(run.get("pm_names", []))

    col_info, col_view, col_dl, col_del = st.columns([5, 1, 1, 1])

    with col_info:
        st.markdown(
            f"**{name}** &nbsp;&nbsp; `{created}` &nbsp;&nbsp; "
            f"{num_p} properties &nbsp;&nbsp; {yrs} &nbsp;&nbsp; *{pms}*"
        )

    with col_view:
        if st.button("View", key=f"view_{run_id}"):
            st.session_state["current_run_id"] = run_id
            st.switch_page("pages/dashboard.py")

    with col_dl:
        # Build ZIP on demand
        run_dir   = os.path.join("runs", run_id)
        meta      = run
        main_wb   = os.path.join(run_dir, meta.get("main_workbook", ""))
        backup_wb = os.path.join(run_dir, meta.get("backup_workbook", ""))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            if os.path.isfile(main_wb):
                zf.write(main_wb, os.path.basename(main_wb))
            if os.path.isfile(backup_wb):
                zf.write(backup_wb, os.path.basename(backup_wb))
        buf.seek(0)
        st.download_button(
            "⬇",
            data=buf,
            file_name=f"{name} Analysis Workbooks.zip",
            mime="application/zip",
            key=f"dl_{run_id}",
        )

    with col_del:
        if st.button("🗑", key=f"del_{run_id}"):
            # Show a confirmation before deleting
            st.session_state[f"confirm_del_{run_id}"] = True

    # Confirmation row
    if st.session_state.get(f"confirm_del_{run_id}"):
        st.warning(f"Delete **{name}** ({created})? This cannot be undone.")
        c1, c2, _ = st.columns([1, 1, 6])
        with c1:
            if st.button("Yes, delete", key=f"yes_{run_id}", type="primary"):
                delete_run(run_id)
                if st.session_state.get("current_run_id") == run_id:
                    del st.session_state["current_run_id"]
                del st.session_state[f"confirm_del_{run_id}"]
                st.rerun()
        with c2:
            if st.button("Cancel", key=f"no_{run_id}"):
                del st.session_state[f"confirm_del_{run_id}"]
                st.rerun()

    st.divider()
```

- [ ] **Step 2: Manually test history page**

```powershell
streamlit run streamlit_app.py
```

- Run at least one analysis
- Navigate to History
- Verify the run appears with the correct date and portfolio name
- Verify "View" navigates to the dashboard with that run loaded
- Verify "⬇" downloads the ZIP
- Verify the delete confirmation flow works

Stop the server.

- [ ] **Step 3: Commit**

```powershell
git add pages/history.py
git commit -m "feat: Streamlit history page (view, download, delete with confirmation)"
```

---

## Task 11: Replace Flask Integration Tests

**Files:**
- Delete: `tests/test_routes.py`
- The replacement `tests/test_pipeline.py` was already created in Task 6.

The `test_routes.py` file tests Flask routes that will be deleted. The pipeline tests in `test_pipeline.py` cover the same logic paths (run analysis, check metadata, check workbooks, validate error handling).

- [ ] **Step 1: Run the new pipeline tests to confirm they're green**

```powershell
pytest tests/test_pipeline.py -v
```

Expected: All 5 PASS.

- [ ] **Step 2: Run the full test suite (excluding the old route tests)**

```powershell
pytest tests/ -q --ignore=tests/test_routes.py
```

Expected: All pass.

- [ ] **Step 3: Delete `tests/test_routes.py`**

```powershell
Remove-Item tests\test_routes.py
```

- [ ] **Step 4: Run the full test suite with no exclusions**

```powershell
pytest tests/ -q
```

Expected: All tests pass (should be more than before, since test_routes.py had Flask-specific imports that could fail).

- [ ] **Step 5: Commit**

```powershell
git add -A
git commit -m "test: replace Flask route tests with pipeline integration tests"
```

---

## Task 12: Delete the Flask Layer

**Files:**
- Delete: `app/routes/` (entire directory)
- Delete: `app/templates/` (entire directory)
- Delete: `app/__init__.py`
- Delete: `app.py`
- Delete: `static/style.css`

Do this ONLY after all tests pass in Task 11. This step is irreversible (though tracked in git).

- [ ] **Step 1: Confirm all tests pass before proceeding**

```powershell
pytest tests/ -q
```

Expected: All pass. If any fail, fix them before continuing.

- [ ] **Step 2: Delete the Flask-specific files and directories**

```powershell
Remove-Item -Recurse -Force app\routes
Remove-Item -Recurse -Force app\templates
Remove-Item app\__init__.py
Remove-Item app.py
Remove-Item static\style.css
```

- [ ] **Step 3: Run tests again to confirm nothing in the test suite imported from the deleted files**

```powershell
pytest tests/ -q
```

Expected: All pass. If any test imports `from app import create_app` or `from app.routes.*`, it will fail here. Fix by removing that import.

- [ ] **Step 4: Verify the Streamlit app still starts**

```powershell
streamlit run streamlit_app.py
```

Expected: App starts at `http://localhost:8501`, all four pages accessible. No import errors.

Stop the server.

- [ ] **Step 5: Final commit**

```powershell
git add -A
git commit -m "feat: remove Flask layer (routes, templates, static CSS, app factory)

Full Streamlit port complete. Backend unchanged. Flask dependency removed."
```

---

## Verification Checklist

After all tasks are complete, verify end-to-end:

- [ ] `pytest tests/ -q` — all tests pass
- [ ] `streamlit run streamlit_app.py` — app starts at `http://localhost:8501`
- [ ] Upload a real financial workbook → analysis runs, dashboard loads with populated KPI table
- [ ] KPI table shows all group rows (Income, Expenses, NOI, GPR, Eco Occ, Per-Unit)
- [ ] Download button on dashboard → ZIP downloads with two valid Excel files (open in Excel without errors)
- [ ] Property detail navigation (dropdown → page switch) works
- [ ] Property detail shows monthly KPIs, full-year projection, AR aging (if AR uploaded)
- [ ] History page shows the run after analysis
- [ ] "View" from history → dashboard reloads with that run
- [ ] Delete from history → run folder removed, no longer appears in history
- [ ] Both workbooks open in Excel without repair prompts
- [ ] No Flask imports remain anywhere in the codebase: `grep -r "from flask\|import flask" app/ pages/ streamlit_app.py` returns nothing
