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
