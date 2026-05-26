# Dashboard KPI UX Improvements — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` (recommended) or `superpowers:executing-plans` to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add collapsible KPI groups to the Portfolio Summary table and reorder the Per Property Analysis table columns to interleave per-unit metrics beside their base metrics.

**Architecture:** Three files change. `app/routes/results.py` extends the flat `_SUMMARY_KPI_DEFINITIONS` list from 4-tuples to 6-tuples, adding `group_id` and `is_group_header`; the builder loop derives `is_group_child` and stamps all three fields onto every summary row dict. `app/templates/dashboard.html` consumes the new fields to render group-header rows (with `data-group` and onclick) and child rows (with `data-parent`, hidden by default via CSS), and includes a small vanilla-JS `toggleKpiGroup` function. `static/style.css` adds the KPI-card header style and the `display:none` rule for child rows. Changes are additive and backward-compatible — the route change doesn't touch the template, so Tasks 1 and 2 are safe to land independently.

**Tech Stack:** Python 3.11 / Flask 3 / Jinja2, vanilla JavaScript (ES5), CSS

---

## File Map

```
app/routes/results.py           — extend _SUMMARY_KPI_DEFINITIONS to 6-tuples; update
                                  builder loop; add divider before Eco Occ %;
                                  remove divider between Eco Occ Variance and Physical Occ %
static/style.css                — add .kpi-group-header, .kpi-group-child,
                                  and chevron transition rules
app/templates/dashboard.html   — (a) property table column reorder
                                  (b) Portfolio Summary tbody: three row branches
                                      (header / child / plain) + <script> toggle block
tests/test_routes.py            — two new tests: group-metadata smoke test +
                                  rendered-HTML data-group assertion
```

---

### Task 1: Extend KPI definitions with group metadata and fix dividers

**Files:**
- Modify: `app/routes/results.py`
- Modify: `tests/test_routes.py`

The `_SUMMARY_KPI_DEFINITIONS` tuples grow from `(label, key, fmt, fav)` to
`(label, key, fmt, fav, group_id, is_group_header)`.
`group_id` is `None` for rows that don't belong to any collapsible group
(Physical Occ %, Leakage Gap, and the three per-unit rows).
The builder loop derives `is_group_child = group_id is not None and not is_group_header`
and adds all three fields to each row dict.

Divider changes (the `None` sentinels in the list):
- A new `None` is inserted between Net Collectible and Eco Occ %.
- The existing `None` between Eco Occ Variance and Physical Occ % is removed.

- [ ] **Step 1: Add the smoke-test to `tests/test_routes.py`**

Append after the last test:

```python
def test_summary_kpi_rows_have_group_metadata(client):
    """Route builds summary_kpi_rows with group fields — no crash on 6-tuple definitions."""
    wb_bytes = _make_workbook_bytes()
    resp = client.post(
        "/",
        data={
            "portfolio_name": "Group Meta Test",
            "eco_occ_target": "95",
            "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert resp.status_code == 200
    assert b"Portfolio Summary" in resp.data
```

- [ ] **Step 2: Run it to confirm it passes right now (baseline)**

```powershell
cd C:\Users\erwin\Desktop\portfolio-analyzer
pytest tests/test_routes.py::test_summary_kpi_rows_have_group_metadata -v
```

Expected: PASS (the current 4-tuple route still works).

- [ ] **Step 3: Replace `_SUMMARY_KPI_DEFINITIONS` in `app/routes/results.py`**

Find the list that starts at line 36 (`_SUMMARY_KPI_DEFINITIONS = [`) and ends at line 67 (closing `]`). Replace the entire list with:

```python
_SUMMARY_KPI_DEFINITIONS = [
    # (label, key, fmt, favorable_positive, group_id, is_group_header)
    ("Actual Income",      "actual_income",        "currency", None,  "group_income",   True),
    ("Budget Income",      "budget_income",        "currency", None,  "group_income",   False),
    ("Income Variance",    "income_variance",      "currency", True,  "group_income",   False),
    ("Income Variance %",  "income_variance_pct",  "pct",      True,  "group_income",   False),
    None,
    ("Actual Expenses",    "actual_expenses",      "currency", None,  "group_expenses", True),
    ("Budget Expenses",    "budget_expenses",      "currency", None,  "group_expenses", False),
    ("Expense Variance",   "expense_variance",     "currency", False, "group_expenses", False),
    ("Expense Variance %", "expense_variance_pct", "pct",      False, "group_expenses", False),
    None,
    ("Actual NOI",         "actual_noi",           "currency", None,  "group_noi",      True),
    ("Budget NOI",         "budget_noi",           "currency", None,  "group_noi",      False),
    ("NOI Variance",       "noi_variance",         "currency", True,  "group_noi",      False),
    ("NOI Variance %",     "noi_variance_pct",     "pct",      True,  "group_noi",      False),
    None,
    ("GPR",                "gpr",                  "currency", None,  "group_gpr",      True),
    ("Vacancy",            "vacancy",              "currency", None,  "group_gpr",      False),
    ("Concessions",        "concessions",          "currency", None,  "group_gpr",      False),
    ("Bad Debt",           "bad_debt",             "currency", None,  "group_gpr",      False),
    ("Net Collectible",    "net_collectible",      "currency", None,  "group_gpr",      False),
    None,                                                              # NEW — divider before Eco Occ %
    ("Eco Occ %",          "eco_occ_pct",          "pct",      None,  "group_eco_occ",  True),
    ("Budget Eco Occ %",   "budget_eco_occ_pct",   "pct",      None,  "group_eco_occ",  False),
    ("Eco Occ Variance",   "eco_occ_variance",     "pct",      True,  "group_eco_occ",  False),
    # divider removed here — Physical Occ % and Leakage Gap join the eco occ section
    ("Physical Occ %",     "physical_occ_pct",     "pct",      None,  None,             False),
    ("Leakage Gap",        "leakage_gap",          "pct",      False, None,             False),
    None,
    ("Income/Unit",        "income_per_unit",      "currency", None,  None,             False),
    ("Expense/Unit",       "expense_per_unit",     "currency", None,  None,             False),
    ("NOI/Unit",           "noi_per_unit",         "currency", None,  None,             False),
]
```

- [ ] **Step 4: Update the `summary_kpi_rows` builder loop in `app/routes/results.py`**

Find the loop that starts with `summary_kpi_rows = []` (around line 279). Replace the entire loop with:

```python
    summary_kpi_rows = []
    for defn in _SUMMARY_KPI_DEFINITIONS:
        if defn is None:
            summary_kpi_rows.append({"sep": True})
            continue
        label, key, fmt, fav, group_id, is_group_header = defn
        is_group_child = group_id is not None and not is_group_header
        values = [period_aggs.get(lbl, {}).get(key) for lbl in period_labels]
        summary_kpi_rows.append({
            "sep":               False,
            "label":             label,
            "key":               key,
            "fmt":               fmt,
            "favorable_positive": fav,
            "tooltip":           _KPI_TOOLTIPS.get(label, ""),
            "period_values":     values,
            "group_id":          group_id,
            "is_group_header":   is_group_header,
            "is_group_child":    is_group_child,
        })
```

- [ ] **Step 5: Run the full test suite**

```powershell
pytest tests/ -q
```

Expected: `85 passed` (84 original + 1 new).

- [ ] **Step 6: Commit**

```powershell
git add app/routes/results.py tests/test_routes.py
git commit -m "feat: extend KPI definitions with group metadata and fix summary dividers"
```

---

### Task 2: CSS — KPI card group-header styles

**Files:**
- Modify: `static/style.css`

Add rules for `.kpi-group-header` (the always-visible clickable header row) and
`.kpi-group-child` (child rows, hidden by default). The chevron is a CSS `::before`
pseudo-element on the `.kpi-label` cell; it rotates 90° when the header has the
`.open` class. The hover background must override the generic `.kpi-table tr:hover td`
rule already in the stylesheet, so the rule is written to be more specific.

- [ ] **Step 1: Append the following block to the end of `static/style.css`**

```css
/* ── Collapsible KPI group rows ───────────────────────────────────────────── */

/* Header row — styled as a clickable KPI card */
.kpi-group-header td {
  background: #ebf3fb;
  font-weight: 600;
  cursor: pointer;
  user-select: none;
}
/* More-specific hover to override .kpi-table tr:hover td */
.kpi-table .kpi-group-header:hover td {
  background: #d6e8f7;
}

/* Chevron in the label cell */
.kpi-group-header .kpi-label::before {
  content: "\25B6\00A0"; /* ▶ + non-breaking space */
  font-size: 9px;
  display: inline-block;
  transition: transform 0.18s ease;
  color: #1f4e79;
  vertical-align: middle;
}
.kpi-group-header.open .kpi-label::before {
  transform: rotate(90deg);
}

/* Child rows — hidden until the group is expanded */
.kpi-group-child {
  display: none;
}
.kpi-group-child td {
  background: #f7fbff;
}
.kpi-group-child .kpi-label {
  padding-left: 28px; /* indent to signal hierarchy */
}
```

- [ ] **Step 2: Run the full test suite (CSS-only change — no Python affected)**

```powershell
pytest tests/ -q
```

Expected: `85 passed`.

- [ ] **Step 3: Commit**

```powershell
git add static/style.css
git commit -m "feat: CSS styles for collapsible KPI group headers and child rows"
```

---

### Task 3: Per Property Analysis — column reorder

**Files:**
- Modify: `app/templates/dashboard.html`

Replace the `<!-- Property Table -->` section. The new column order is:
Property → PM → Units → Income → Income/Unit → Expenses → Expense/Unit →
NOI → NOI/Unit → NOI Var → Eco Occ % → Phys Occ % → Leakage → Detail.

All fields already exist on every `prop_rows` dict (populated by `_agg_kpis`).
`total_units` is an integer or `None`; render `—` when `None` (same as the `| currency`
filter already does for numeric fields).

- [ ] **Step 1: Replace the property table section in `app/templates/dashboard.html`**

Find the block that starts with `<!-- Property Table -->` and ends with the closing
`</section>` tag before `<!-- AR Aging -->`. Replace the entire block with:

```html
<!-- Property Table -->
<section class="card">
  <h2>Per Property Analysis ({{ latest_period_label }})</h2>
  <div class="table-scroll">
  <table class="data-table">
    <thead>
      <tr>
        <th>Property</th>
        <th>PM</th>
        <th title="Total Units from Physical Occupancy Report">Units</th>
        <th title="Actual Income = GPR + Other Income − Vacancy − Concessions − Bad Debt">Income</th>
        <th title="Actual Income / Total Units">Income/Unit</th>
        <th title="Total Operating Expenses">Expenses</th>
        <th title="Actual Expenses / Total Units">Expense/Unit</th>
        <th title="NOI = Income − Expenses">NOI</th>
        <th title="NOI / Total Units">NOI/Unit</th>
        <th title="NOI Variance = Actual − Budget">NOI Var</th>
        <th title="Economic Occ % = Net Collectible / GPR">Eco Occ %</th>
        <th title="Physical Occ % from occupancy report">Phys Occ %</th>
        <th title="Physical Occ % − Economic Occ %">Leakage</th>
        <th>Detail</th>
      </tr>
    </thead>
    <tbody>
      {% for r in prop_rows %}
      <tr>
        <td>{{ r.property_name }}</td>
        <td>{{ r.pm_name }}</td>
        <td class="currency">{{ r.total_units if r.total_units is not none else '—' }}</td>
        <td class="currency">{{ r.actual_income | currency }}</td>
        <td class="currency">{{ r.income_per_unit | currency }}</td>
        <td class="currency">{{ r.actual_expenses | currency }}</td>
        <td class="currency">{{ r.expense_per_unit | currency }}</td>
        <td class="currency">{{ r.actual_noi | currency }}</td>
        <td class="currency">{{ r.noi_per_unit | currency }}</td>
        <td class="currency {{ 'pos' if (r.noi_variance or 0) >= 0 else 'neg' }}">{{ r.noi_variance | currency }}</td>
        <td class="pct {% if r.eco_occ_pct is not none and r.eco_occ_pct < eco_occ_target %}neg{% endif %}">{{ r.eco_occ_pct | pct }}</td>
        <td class="pct">{{ r.physical_occ_pct | pct }}</td>
        <td class="pct {{ 'neg' if (r.leakage_gap or 0) > 0 else '' }}">{{ r.leakage_gap | pct }}</td>
        <td><a href="{{ url_for('property_detail.show', run_id=run_id, property_name=r.property_name) }}">View</a></td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
  </div>
</section>
```

- [ ] **Step 2: Run the full test suite**

```powershell
pytest tests/ -q
```

Expected: `85 passed`.

- [ ] **Step 3: Commit**

```powershell
git add app/templates/dashboard.html
git commit -m "feat: reorder property table columns — interleave per-unit metrics beside base metrics"
```

---

### Task 4: Portfolio Summary — collapsible group rows and JS toggle

**Files:**
- Modify: `app/templates/dashboard.html`
- Modify: `tests/test_routes.py`

The `<tbody>` of the Portfolio Summary table gains three rendering branches instead of one:

1. **`row.is_group_header`** → `<tr class="kpi-group-header" data-group="…" onclick="…">`
2. **`row.is_group_child`** → `<tr class="kpi-group-child" data-parent="…">` (hidden by CSS)
3. **plain row** → `<tr>` (unchanged from today)

The value-cell markup inside all three branches is identical, so a Jinja2
`{% macro %}` is defined once above the table and called in each branch to avoid
repetition.

A `<script>` block is appended before `{% endblock %}` containing the single
`toggleKpiGroup` function. It uses `row.style.display = 'table-row'` (not `''`)
to show children, because the CSS class `.kpi-group-child { display: none }` would
win over an empty inline style.

- [ ] **Step 1: Add the HTML assertion test to `tests/test_routes.py`**

Append after the last test:

```python
def test_dashboard_has_collapsible_group_attributes(client):
    """Rendered dashboard contains data-group / data-parent attributes for all five groups."""
    wb_bytes = _make_workbook_bytes()
    resp = client.post(
        "/",
        data={
            "portfolio_name": "Collapse HTML Test",
            "eco_occ_target": "95",
            "financial_files": (io.BytesIO(wb_bytes), "test_financial.xlsx"),
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert resp.status_code == 200
    html = resp.data.decode()
    for group in ("group_income", "group_expenses", "group_noi",
                  "group_gpr", "group_eco_occ"):
        assert f'data-group="{group}"'  in html, f"missing data-group={group}"
        assert f'data-parent="{group}"' in html, f"missing data-parent={group}"
    assert 'class="kpi-group-child"' in html
    assert "toggleKpiGroup" in html
```

- [ ] **Step 2: Run it to confirm it currently fails**

```powershell
pytest tests/test_routes.py::test_dashboard_has_collapsible_group_attributes -v
```

Expected: FAIL — `AssertionError: missing data-group=group_income`
(the template doesn't emit those attributes yet).

- [ ] **Step 3: Replace the Portfolio Summary `<tbody>` in `app/templates/dashboard.html`**

Find the entire `<section class="card">` block that starts with
`<!-- Portfolio Summary` and ends with its closing `</section>`.
Replace the `<table class="kpi-table kpi-transposed">` block (keeping the `<section>`
and `<h2>` wrapper intact) with the version below.

The key change is the `<tbody>`: a macro is declared just before the `<table>` so
that value-cell rendering logic is written once, and the `<tbody>` uses `{% if %}`
branches to emit the correct row class and attributes.

```html
<!-- Portfolio Summary — transposed: KPI labels as rows, quarter-periods as columns -->
<section class="card">
  <h2>Portfolio Summary</h2>
  <div class="table-scroll">

  {# Macro: renders the period-value <td> cells for any row type #}
  {% macro value_cells(row) %}
    {% for val in row.period_values %}
    {%- set fav = row.favorable_positive -%}
    {%- if fav is not none -%}
      {%- if fav -%}
        {%- set color_cls = 'pos' if (val or 0) >= 0 else 'neg' -%}
      {%- else -%}
        {%- set color_cls = 'neg' if (val or 0) > 0 else ('pos' if (val or 0) < 0 else '') -%}
      {%- endif -%}
    {%- else -%}
      {%- set color_cls = '' -%}
    {%- endif -%}
    <td class="{% if row.fmt == 'currency' %}currency{% else %}pct{% endif %} {{ color_cls }}">
      {%- if row.fmt == 'currency' -%}{{ val | currency }}{%- else -%}{{ val | pct }}{%- endif -%}
    </td>
    {% endfor %}
  {% endmacro %}

  <table class="kpi-table kpi-transposed">
    <thead>
      <tr>
        <th class="kpi-label-col">KPI</th>
        {% for lbl in period_labels %}
        <th>{{ lbl }}<br><span class="period-prop-count">{{ period_property_counts.get(lbl, 0) }} Properties</span></th>
        {% endfor %}
      </tr>
    </thead>
    <tbody>
      {% for row in summary_kpi_rows %}
      {% if row.sep %}
      <tr class="kpi-sep"><td colspan="{{ period_labels|length + 1 }}"></td></tr>

      {% elif row.is_group_header %}
      <tr class="kpi-group-header"
          data-group="{{ row.group_id }}"
          onclick="toggleKpiGroup('{{ row.group_id }}', this)">
        <td class="kpi-label" title="{{ row.tooltip }}">{{ row.label }}</td>
        {{ value_cells(row) }}
      </tr>

      {% elif row.is_group_child %}
      <tr class="kpi-group-child" data-parent="{{ row.group_id }}">
        <td class="kpi-label" title="{{ row.tooltip }}">{{ row.label }}</td>
        {{ value_cells(row) }}
      </tr>

      {% else %}
      <tr>
        <td class="kpi-label" title="{{ row.tooltip }}">{{ row.label }}</td>
        {{ value_cells(row) }}
      </tr>
      {% endif %}
      {% endfor %}
    </tbody>
  </table>
  </div>
</section>
```

- [ ] **Step 4: Add the `<script>` block before `{% endblock %}` at the very end of `dashboard.html`**

The last two lines of the file are currently:
```
{% endif %}
{% endblock %}
```

Insert the script block between them so the file ends with:

```html
<script>
function toggleKpiGroup(groupId, headerRow) {
  var children = document.querySelectorAll('[data-parent="' + groupId + '"]');
  var isOpen = headerRow.classList.contains('open');
  if (isOpen) {
    headerRow.classList.remove('open');
    /* Remove inline style — CSS class .kpi-group-child { display:none } takes over */
    children.forEach(function(row) { row.style.display = ''; });
  } else {
    headerRow.classList.add('open');
    /* Inline style overrides the CSS class display:none */
    children.forEach(function(row) { row.style.display = 'table-row'; });
  }
}
</script>
{% endblock %}
```

- [ ] **Step 5: Run the new test**

```powershell
pytest tests/test_routes.py::test_dashboard_has_collapsible_group_attributes -v
```

Expected: PASS.

- [ ] **Step 6: Run the full suite**

```powershell
pytest tests/ -q
```

Expected: `87 passed` (85 from Tasks 1–3 + 2 new tests from Tasks 1 and 4).

- [ ] **Step 7: Commit**

```powershell
git add app/templates/dashboard.html tests/test_routes.py
git commit -m "feat: collapsible KPI groups in Portfolio Summary with JS expand/collapse"
```

---

## Verification Checklist

After all four tasks are complete:

- [ ] `pytest tests/ -q` → `87 passed`, 0 failures
- [ ] `python app.py` starts without errors at `http://localhost:5000`
- [ ] Open any existing run's dashboard. On page load the Portfolio Summary shows **only** the five group-header rows (Actual Income, Actual Expenses, Actual NOI, GPR, Eco Occ %), Physical Occ %, Leakage Gap, and the three per-unit rows — all child rows are hidden
- [ ] A visible section divider (grey band) appears between Net Collectible and Eco Occ %
- [ ] No section divider appears between Eco Occ Variance and Physical Occ %
- [ ] Clicking **Actual Income** reveals Budget Income / Income Variance / Income Variance % and the chevron rotates; clicking again collapses them
- [ ] Same expand/collapse behaviour for **Actual Expenses**, **Actual NOI**, **GPR**, and **Eco Occ %**
- [ ] Child rows are visually indented (left padding) and on a slightly lighter blue background
- [ ] Per Property Analysis table column order: Property · PM · Units · Income · Income/Unit · Expenses · Expense/Unit · NOI · NOI/Unit · NOI Var · Eco Occ % · Phys Occ % · Leakage · Detail
- [ ] Units column shows an integer when a physical occupancy report was uploaded, `—` when not
- [ ] Income/Unit, Expense/Unit, NOI/Unit show `—` when no occupancy report was uploaded (same as existing `| currency` filter behaviour for `None`)
