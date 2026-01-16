import sys
import os
sys.path.append(os.path.dirname(__file__))

import streamlit as st
import pandas as pd
from datetime import date, datetime, time

from spike_backend import fetch_incidents_for_range, generate_excel, IST
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
h1 { margin-bottom: 0.2rem !important; font-size: 28px !important; }
h2, h3 { margin-top: 0.2rem !important; margin-bottom: 0.3rem !important; }
.stButton>button { padding: 5px 18px; font-size: 14px; border-radius: 10px; }
div[data-testid="column"] input[type="date"] { max-width: 150px; padding: 2px 4px; font-size: 12px; }
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
    st.subheader("Incident Report (Date & Time Range)")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        from_date = st.date_input("From Date", value=date.today())
    with col2:
        from_time = st.time_input("From Time", value=time(0, 0))
    with col3:
        to_date = st.date_input("To Date", value=date.today())
    with col4:
        to_time = st.time_input("To Time", value=time(23, 59))

    if st.button("Fetch Report"):
        # Make user-selected datetimes aware (IST)
        from_dt = datetime.combine(from_date, from_time).replace(tzinfo=IST)
        to_dt = datetime.combine(to_date, to_time).replace(tzinfo=IST)

        if from_dt > to_dt:
            st.error("From datetime cannot be after To datetime")
        else:
            with st.spinner("Fetching incidents..."):
                rows = fetch_incidents_for_range(from_dt, to_dt)

            if not rows:
                st.warning("No incidents found")
            else:
                df = pd.DataFrame(rows)
                df.sort_values(by="Created (IST)", ascending=False, inplace=True)
                st.success(f"Total Incidents: {len(df)}")
                st.dataframe(df, use_container_width=True)

                file = generate_excel(rows, from_dt, to_dt)
                st.download_button(
                    "📥 Download Excel",
                    data=open(file, "rb"),
                    file_name=file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# -------------------------------------------------
# OPEN ALERTS
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
            df.sort_values(by="Created (IST)", ascending=False, inplace=True)
            st.success(f"Total Open Alerts: {len(df)}")
            st.dataframe(df, use_container_width=True)

            file = open_generate_excel(rows)
            st.download_button(
                "📥 Download Open Alerts Excel",
                data=open(file, "rb"),
                file_name=file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
