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
st.set_page_config(page_title="Spike NOC Dashboard", layout="wide")

# -------------------------------------------------
# SESSION STATE INIT
# -------------------------------------------------
defaults = {
    "from_date": date.today(),
    "from_time": time(0, 0),
    "to_date": date.today(),
    "to_time": time(23, 59),
    "incident_rows": None,
    "open_alert_rows": None,
    "incident_selected_teams": ["All Teams"],
    "open_selected_teams": ["All Teams"],
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# -------------------------------------------------
# SIDEBAR
# -------------------------------------------------
st.sidebar.image("logo.png", width=180)
st.sidebar.title("📌 Navigation")

page = st.sidebar.radio("Select View", ["Incident Report", "Open Alerts"])

st.title("Spike NOC Dashboard")

# =================================================
# INCIDENT REPORT
# =================================================
if page == "Incident Report":

    st.subheader("📄 Incident Report (Date, Time & Team)")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.session_state.from_date = st.date_input(
            "From Date", value=st.session_state.from_date
        )

    with c2:
        st.session_state.from_time = st.time_input(
            "From Time", value=st.session_state.from_time
        )

    with c3:
        st.session_state.to_date = st.date_input(
            "To Date", value=st.session_state.to_date
        )

    with c4:
        st.session_state.to_time = st.time_input(
            "To Time", value=st.session_state.to_time
        )

    if st.button("Fetch Report"):
        from_dt = datetime.combine(
            st.session_state.from_date,
            st.session_state.from_time
        ).replace(tzinfo=IST)

        to_dt = datetime.combine(
            st.session_state.to_date,
            st.session_state.to_time
        ).replace(tzinfo=IST)

        if from_dt > to_dt:
            st.error("From datetime cannot be after To datetime")
        else:
            with st.spinner("Fetching incidents..."):
                st.session_state.incident_rows = fetch_incidents_for_range(from_dt, to_dt)

    # ---------------- DISPLAY DATA ----------------
    if st.session_state.incident_rows is not None:

        if not st.session_state.incident_rows:
            st.warning("No incidents found")
        else:
            df = pd.DataFrame(st.session_state.incident_rows)

            # Team selector
            teams = sorted(df["Team Name"].dropna().unique().tolist())
            team_options = ["All Teams"] + teams

            st.session_state.incident_selected_teams = st.multiselect(
                "Select Team(s)",
                team_options,
                default=st.session_state.incident_selected_teams
            )

            # 🚫 No team selected
            if not st.session_state.incident_selected_teams:
                st.warning("⚠️ Please select at least one team")
                st.stop()

            # Apply filter
            if "All Teams" not in st.session_state.incident_selected_teams:
                df = df[df["Team Name"].isin(st.session_state.incident_selected_teams)]

            # Sort by latest Created (IST)
            if "Created (IST)" in df.columns:
                df["Created (IST)"] = pd.to_datetime(df["Created (IST)"])
                df = df.sort_values("Created (IST)", ascending=False)

            df.reset_index(drop=True, inplace=True)

            st.success(f"Total Incidents: {len(df)}")
            st.dataframe(df, use_container_width=True, hide_index=True)

            # Excel matches UI
            excel_rows = df.to_dict(orient="records")

            file = generate_excel(
                excel_rows,
                datetime.combine(st.session_state.from_date, st.session_state.from_time).replace(tzinfo=IST),
                datetime.combine(st.session_state.to_date, st.session_state.to_time).replace(tzinfo=IST)
            )

            st.download_button(
                "📥 Download Excel",
                data=open(file, "rb"),
                file_name=file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =================================================
# OPEN ALERTS
# =================================================
if page == "Open Alerts":

    st.subheader("🚨 Open Alerts (All Teams / Per Team)")

    if st.button("Refresh Open Alerts"):
        with st.spinner("Fetching open alerts..."):
            st.session_state.open_alert_rows = fetch_all_open_alerts()

    if st.session_state.open_alert_rows is not None:

        if not st.session_state.open_alert_rows:
            st.success("🎉 No open alerts found")
        else:
            df = pd.DataFrame(st.session_state.open_alert_rows)

            teams = sorted(df["Team Name"].dropna().unique().tolist())
            team_options = ["All Teams"] + teams

            st.session_state.open_selected_teams = st.multiselect(
                "Select Team(s)",
                team_options,
                default=st.session_state.open_selected_teams
            )

            # 🚫 No team selected
            if not st.session_state.open_selected_teams:
                st.warning("⚠️ Please select at least one team")
                st.stop()

            if "All Teams" not in st.session_state.open_selected_teams:
                df = df[df["Team Name"].isin(st.session_state.open_selected_teams)]

            # Sort by latest Created (IST)
            if "Created (IST)" in df.columns:
                df["Created (IST)"] = pd.to_datetime(df["Created (IST)"])
                df = df.sort_values("Created (IST)", ascending=False)

            df.reset_index(drop=True, inplace=True)

            st.success(f"Total Open Alerts: {len(df)}")
            st.dataframe(df, use_container_width=True, hide_index=True)

            file = open_generate_excel(df.to_dict(orient="records"))

            st.download_button(
                "📥 Download Open Alerts Excel",
                data=open(file, "rb"),
                file_name=file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
