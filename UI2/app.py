import sys
import os
sys.path.append(os.path.dirname(__file__))

import streamlit as st
import pandas as pd
from datetime import date

from spike_backend import fetch_incidents_for_range, generate_excel
from open_alerts_backend import fetch_all_open_alerts, generate_excel as open_generate_excel

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Spike NOC Dashboard",
    layout="wide"
)

# -------------------------------------------------
# GLOBAL CSS
# -------------------------------------------------
st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem !important;
    padding-left: 2rem;
    padding-right: 2rem;
}

h1 {
    margin-bottom: 0.2rem !important;
    font-size: 28px !important;
}

h2, h3 {
    margin-top: 0.2rem !important;
    margin-bottom: 0.3rem !important;
}

.stButton>button {
    padding: 5px 18px;
    font-size: 14px;
    border-radius: 10px;
}

div[data-testid="column"] input[type="date"] {
    max-width: 150px;
    padding: 2px 4px;
    font-size: 12px;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# SIDEBAR
# -------------------------------------------------
st.sidebar.image("logo.png", width=180)
st.sidebar.title("📌 Navigation")

page = st.sidebar.radio(
    "Select View",
    ["Incident Report", "Open Alerts"]
)

# -------------------------------------------------
# TITLE
# -------------------------------------------------
st.title("Spike NOC Dashboard")

# -------------------------------------------------
# INCIDENT REPORT
# -------------------------------------------------
if page == "Incident Report":

    st.subheader("Incident Report (Date Range)")

    col1, col2 = st.columns(2)
    with col1:
        from_date = st.date_input("From Date", value=date.today())
    with col2:
        to_date = st.date_input("To Date", value=date.today())

    if st.button("Fetch Report"):
        with st.spinner("Fetching incidents..."):
            rows = fetch_incidents_for_range(from_date, to_date)

        if not rows:
            st.warning("No incidents found")
        else:
            df = pd.DataFrame(rows)
            st.success(f"Total Incidents: {len(df)}")
            st.dataframe(df, use_container_width=True)

            file = generate_excel(rows, from_date, to_date)
            st.download_button(
                "📥 Download Excel",
                data=open(file, "rb"),
                file_name=file
            )

# -------------------------------------------------
# OPEN ALERTS (USING YOUR LOGIC)
# -------------------------------------------------
if page == "Open Alerts":

    st.subheader("🚨 Open Alerts (All Teams)")

    if st.button("Refresh Open Alerts"):
        with st.spinner("Fetching open alerts..."):
            rows = fetch_all_open_alerts()

        if not rows:
            st.success("🎉 No open alerts found")
        else:
            df = pd.DataFrame(rows)
            st.success(f"Total Open Alerts: {len(df)}")
            st.dataframe(df, use_container_width=True)

            file = open_generate_excel(rows)
            st.download_button(
                "📥 Download Open Alerts Excel",
                data=open(file, "rb"),
                file_name=file
            )
