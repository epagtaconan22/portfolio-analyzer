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
