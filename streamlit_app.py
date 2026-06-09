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
