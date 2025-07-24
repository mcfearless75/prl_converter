import os
import streamlit as st
import pandas as pd
import docx
import pdfplumber
import re
import unicodedata
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ==== DB Connection: Use Postgres on Render, SQLite locally ====
if "DATABASE_URL" in os.environ:
    import psycopg2
    from urllib.parse import urlparse
    url = urlparse(os.environ["DATABASE_URL"])
    conn = psycopg2.connect(
        dbname=url.path[1:],
        user=url.username,
        password=url.password,
        host=url.hostname,
        port=url.port
    )
else:
    import sqlite3
    conn = sqlite3.connect("timesheets.db", check_same_thread=False)

c = conn.cursor()
c.execute("""
CREATE TABLE IF NOT EXISTS timesheet_entries (
    id SERIAL PRIMARY KEY,
    name TEXT,
    matched_as TEXT,
    ratio REAL,
    client TEXT,
    site_address TEXT,
    department TEXT,
    weekday_hours REAL,
    saturday_hours REAL,
    sunday_hours REAL,
    rate REAL,
    date_range TEXT,
    extracted_on TEXT,
    source_file TEXT,
    upload_timestamp TIMESTAMP
)
""")

# ==== Helper Functions & Rate DB ====
RATE_FILE_PATH = "pay_rates.xlsx"

def normalize_name(name: str) -> str:
    nfkd = unicodedata.normalize("NFKD", name)
    only_ascii = nfkd.encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z]", "", only_ascii).lower()

def load_rate_database(excel_path: str):
    custom_rates = {}
    normalized_rates = {}
    norm_to_raw = {}
    wb = load_workbook(excel_path, data_only=True)
    for sheet in wb.sheetnames:
        df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        header_row = None
        for idx, val in enumerate(df_raw[0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                second = df_raw.iat[idx, 1]
                if isinstance(second, str) and second.strip().lower() == "pay rate":
                    header_row = idx
                    break
        if header_row is None:
            continue
        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
        if "Name" not in df.columns or "Pay Rate" not in df.columns:
            continue
        df = df[["Name", "Pay Rate"]].copy()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df = df.dropna(subset=["Name", "Pay Rate"])
        for _, row in df.iterrows():
            raw_name = str(row["Name"]).strip()
            rate = float(row["Pay Rate"])
            custom_rates[raw_name] = rate
            norm = normalize_name(raw_name)
            normalized_rates[norm] = rate
            norm_to_raw[norm] = raw_name
    return custom_rates, normalized_rates, norm_to_raw

custom_rates, normalized_rates, norm_to_raw = load_rate_database(RATE_FILE_PATH)

def lookup_match(name: str):
    norm = normalize_name(name)
    if norm in normalized_rates:
        return norm_to_raw[norm], normalized_rates[norm], 1.0
    else:
        return name, 15.0, 0.0  # default

# (… your extract_from_docx, extract_from_pdf, hhmm_to_hours, calculate_pay, etc. …)

# ====== Streamlit Tabs UI ======
if st.sidebar.button("🔄 Reload Pay Rates"):
    st.cache_data.clear()
    st.experimental_rerun()

st.sidebar.header("Upload Timesheets")
st.sidebar.markdown("""
1. Upload **.docx** or **.pdf**  
2. Confirm name‑matches (expand Debug)  
3. Export Excel with formulas  
""")
uploaded_files = st.sidebar.file_uploader("Choose files", accept_multiple_files=True)

tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1. Upload & Review ----
with tabs[0]:
    st.header("📤 Upload & Review Timesheets")
    # … your existing upload, parse, debug, export logic …

# ---- 2. History ----
with tabs[1]:
    import datetime
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Filter by upload date, then see a weekly summary:")

    # 1️⃣ Date picker for upload‑date filtering
    today = datetime.date.today()
    thirty_days_ago = today - datetime.timedelta(days=30)
    start_date, end_date = st.date_input(
        "Select upload date range",
        value=(thirty_days_ago, today),
        min_value=datetime.date(2020, 1, 1),
        max_value=today
    )

    # 2️⃣ Fetch all history rows
    query = """
        SELECT
            name,
            matched_as,
            ratio,
            client,
            site_address,
            department,
            weekday_hours,
            saturday_hours,
            sunday_hours,
            rate AS rate,
            date_range,
            extracted_on,
            source_file,
            upload_timestamp
        FROM timesheet_entries
        ORDER BY upload_timestamp DESC
    """
    c.execute(query)
    history_rows = c.fetchall()
    history_df = pd.DataFrame(history_rows, columns=[
        "Name", "Matched As", "Ratio", "Client", "Site Address", "Department",
        "Weekday Hours", "Saturday Hours", "Sunday Hours", "Rate (£)",
        "Date Range", "Extracted On", "Source File", "Upload Timestamp"
    ])

    # 3️⃣ Filter by the picked date range
    history_df["Upload Timestamp"] = pd.to_datetime(history_df["Upload Timestamp"]).dt.date
    mask = (
        (history_df["Upload Timestamp"] >= start_date)
        & (history_df["Upload Timestamp"] <= end_date)
    )
    filtered = history_df.loc[mask].copy()

    if filtered.empty:
        st.info("No entries in this date range.")
    else:
        # 4️⃣ Compute pay per row
        filtered["Pay"] = (
            filtered["Weekday Hours"]
            + filtered["Saturday Hours"]
            + filtered["Sunday Hours"]
        ) * filtered["Rate (£)"]

        # 5️⃣ Group into weekly summaries
        summary = (
            filtered
            .groupby("Date Range")
            .agg(
                Entries=("Name", "count"),
                Weekday_Hours=("Weekday Hours", "sum"),
                Sat_Hours=("Saturday Hours", "sum"),
                Sun_Hours=("Sunday Hours", "sum"),
                Total_Pay=("Pay", "sum")
            )
            .reset_index()
            .sort_values("Date Range", ascending=False)
        )
        st.markdown("### 📅 Weekly Summary")
        st.dataframe(summary, use_container_width=True)

        # 6️⃣ Raw‑data expander
        with st.expander(f"Show raw entries ({len(filtered)}) rows"):
            st.dataframe(filtered.drop(columns=["Pay"]), use_container_width=True)

# ---- 3. Dashboard ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.markdown("Aggregate stats for all stored timesheets.")
    # … your existing dashboard logic …

# ---- 4. Settings ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Click **Reload Pay Rates** in the sidebar to update rates from your Excel.
    - Only exact matches (case and accent-insensitive) are used for rates.
    - Any unmatched name uses the default rate (£15/hr) and is highlighted in red.
    - For support or feature requests, contact your dev team!
    """)
