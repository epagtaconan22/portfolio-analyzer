# Dashboard KPI UX Improvements — Design Spec

_Date: 2026-05-26_

---

## Goal

Two targeted improvements to the dashboard:

1. **Per Property Analysis table** — add Total Units column and interleave per-unit metrics beside their base metrics.
2. **Portfolio Summary table** — four collapsible KPI groups (collapsed by default, click to expand) with matching divider reorganisation.

---

## Change 1: Per Property Analysis — Column Order

### New column sequence (left → right)

| # | Header | Source field on `prop_rows` |
|---|---|---|
| 1 | Property | `property_name` |
| 2 | PM | `pm_name` |
| 3 | Total Units | `total_units` |
| 4 | Income | `actual_income` |
| 5 | Income/Unit | `income_per_unit` |
| 6 | Expenses | `actual_expenses` |
| 7 | Expense/Unit | `expense_per_unit` |
| 8 | NOI | `actual_noi` |
| 9 | NOI/Unit | `noi_per_unit` |
| 10 | NOI Var | `noi_variance` |
| 11 | Eco Occ % | `eco_occ_pct` |
| 12 | Phys Occ % | `physical_occ_pct` |
| 13 | Leakage | `leakage_gap` |
| 14 | Detail | (link) |

All fields already exist on `prop_rows`. This is a template-only reorder plus one new header.

Formatting rules (unchanged from today):
- Currency columns: use `| currency` filter
- Percentage columns: use `| pct` filter
- NOI Var: `pos` class if ≥ 0, `neg` if < 0
- Eco Occ %: `neg` class if below `eco_occ_target`
- Leakage: `neg` class if > 0
- Total Units / Income/Unit / Expense/Unit / NOI/Unit: render `—` when `None`

---

## Change 2: Portfolio Summary — Collapsible KPI Groups

### Collapsible groups

Four KPI rows become **group headers**. Their detail rows are **children** — hidden by default, shown when the user clicks the header.

| Group header | Children |
|---|---|
| Actual Income | Budget Income · Income Variance · Income Variance % |
| Actual Expenses | Budget Expenses · Expense Variance · Expense Variance % |
| Actual NOI | Budget NOI · NOI Variance · NOI Variance % |
| GPR | Vacancy · Concessions · Bad Debt · Net Collectible |
| Eco Occ % | Budget Eco Occ % · Eco Occ Variance |

Physical Occ % and Leakage Gap are **not** collapsible — they remain always-visible rows in the same section as Eco Occ %.

### Section dividers after changes

```
Actual Income  (group header)
  └ Budget Income, Income Variance, Income Variance %
───────────────────────────────────────────────── (divider — existing)
Actual Expenses  (group header)
  └ Budget Expenses, Expense Variance, Expense Variance %
───────────────────────────────────────────────── (divider — existing)
Actual NOI  (group header)
  └ Budget NOI, NOI Variance, NOI Variance %
───────────────────────────────────────────────── (divider — existing, before GPR)
GPR  (group header)
  └ Vacancy, Concessions, Bad Debt, Net Collectible
───────────────────────────────────────────────── (divider — NEW, before Eco Occ %)
Eco Occ %  (group header)
  └ Budget Eco Occ %, Eco Occ Variance
Physical Occ %                                    ← no divider here (removed)
Leakage Gap
───────────────────────────────────────────────── (divider — existing, before Income/Unit)
Income/Unit
Expense/Unit
NOI/Unit
```

### Visual design — group header rows

Group header rows are styled as clickable KPI cards:

- Background: `#EBF3FB` (light blue, distinct from plain rows)
- Font weight: `600` (semi-bold)
- Cursor: `pointer`
- Left padding: includes a `▶` / `▼` chevron (CSS `::before` pseudo-element or inline span)
- Chevron rotates 90° (CSS transition) when the group is expanded
- Hover: background darkens slightly to `#D6E8F7`
- No red/green coloring on the header row itself (children keep their coloring)

### Behaviour

- All groups start **collapsed** (children hidden) on page load
- Click toggles the group open/closed
- Toggle is implemented with vanilla JavaScript (no dependencies)
- State is not persisted across page loads (always starts collapsed)

---

## Files Changed

| File | Change |
|---|---|
| `app/routes/results.py` | Extend `_SUMMARY_KPI_DEFINITIONS` tuples with `group_id` and `is_group_header` fields. Update `summary_kpi_rows` builder to pass group metadata. Add `None` divider before Eco Occ %; remove `None` divider between Eco Occ Variance and Physical Occ %. |
| `app/templates/dashboard.html` | (1) Reorder property table columns. (2) Render group-header rows with `data-group` attribute and chevron. Render child rows with `data-parent` attribute, `display:none` by default. Add `<script>` toggle block. |
| `static/style.css` | Add `.kpi-group-header` styles (background, cursor, font-weight, hover). Add `.kpi-group-header--open .chevron` rotation rule. Add `.kpi-group-child` rule (initial `display:none`). |

No route logic changes, no model changes, no test changes required (existing tests don't assert on the HTML row order or collapsibility).

---

## Out of Scope

- Persisting expand/collapse state in localStorage
- Collapsing the AR Aging section
- Any changes to the Excel workbook exports
