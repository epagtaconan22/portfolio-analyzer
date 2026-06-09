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
