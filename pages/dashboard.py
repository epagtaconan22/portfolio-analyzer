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
