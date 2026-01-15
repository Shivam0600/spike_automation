import sys
import os
sys.path.append(os.path.dirname(__file__))

import streamlit as st
import pandas as pd
from datetime import date
from spike_backend import fetch_incidents_for_range, generate_excel

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Spike NOC Dashboard",
    layout="wide"
)

# -------------------------------------------------
# CUSTOM CSS FOR STYLISH UI
# -------------------------------------------------
st.markdown("""
    <style>
        /* Block container padding */
        .block-container {
            padding-top: 1rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }

        /* Hide default Streamlit header */
        header {
            visibility: hidden;
            height: 0px;
        }

        /* Center columns */
        .stColumn {
            display: flex;
            justify-content: center;
            align-items: center;
        }

        /* Header row style */
        .header-row {
            background: linear-gradient(90deg, #4facfe 0%, #00f2fe 100%);
            padding: 15px 20px;
            border-radius: 12px;
            align-items: center;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2);
            margin-bottom: 20px;
        }

        .header-row img {
            border-radius: 8px;
        }

        .header-title {
            font-size: 28px;
            font-weight: 700;
            margin-left: 12px;
            color: black; /* text black */
        }

        /* Buttons */
        .stButton>button {
            border-radius: 12px;
            font-weight: 600;
            background: #00b894;
            color: white;
            padding: 6px 20px;
        }

        .stButton>button:hover {
            background-color: #019875;
            color: white;
        }

        /* Download button */
        .stDownloadButton>button {
            border-radius: 12px;
            font-weight: 600;
            background: #0984e3;
            color: white;
            padding: 6px 20px;
        }

        .stDownloadButton>button:hover {
            background-color: #0652dd;
            color: white;
        }

        /* Date inputs style */
        div[data-testid="column"] input[type="date"] {
            max-width: 220px;
            border-radius: 8px;
            border: 1px solid #ccc;
            padding: 5px 8px;
        }

        /* Dataframe style */
        .stDataFrame {
            font-size: 14px;
        }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# HEADER ROW (LOGO + TITLE)
# -------------------------------------------------
header_cols = st.columns([0.5, 2])

with header_cols[0]:
    st.image("logo.png", width=140)

with header_cols[1]:
    st.markdown('<div class="header-title">📊 Spike NOC Incident Report</div>', unsafe_allow_html=True)

# -------------------------------------------------
# FILTER ROW (DATES + FETCH BUTTON)
# -------------------------------------------------
filter_cols = st.columns([1, 1, 0.5, 1])

with filter_cols[0]:
    from_date = st.date_input("From Date", value=date.today())

with filter_cols[1]:
    to_date = st.date_input("To Date", value=date.today())

# with filter_cols[2]:
#     st.write("")  # spacer

with filter_cols[3]:
    fetch_button = st.button("Fetch Report")

# -------------------------------------------------
# FETCH LOGIC
# -------------------------------------------------
if fetch_button:
    if from_date > to_date:
        st.error("From date cannot be after To date")
    else:
        with st.spinner("Fetching incidents from Spike..."):
            rows = fetch_incidents_for_range(from_date, to_date)

        if not rows:
            st.warning("No incidents found for selected date range")
        else:
            df = pd.DataFrame(rows)

            st.success(f"Total Incidents: {len(df)}")
            st.dataframe(df, use_container_width=True)

            # Generate Excel
            excel_file = generate_excel(rows, from_date, to_date)

            # Download button
            st.download_button(
                label="📥 Download Excel Report",
                data=open(excel_file, "rb"),
                file_name=excel_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
