import os
import re
import unicodedata
from pathlib import Path
from io import BytesIO
from datetime import datetime, date, timedelta

import streamlit as st
import pandas as pd

# For DOCX/PDF parsing (your existing code uses these)
import docx
import pdfplumber

# For XLSX rate file
from openpyxl import load_workbook

# ==== DB Connection: Postgres on Render, SQLite locally ====
if "DATABASE_URL" in os.environ:
    import psycopg2
    from urllib.parse import urlparse
    url = urlparse(os.environ["DATABASE_URL"])
    conn = psycopg2.connect(
        dbname=url.path[1:], user=url.username,
        password=url.password, host=url.hostname, port=url.port
    )
else:
    import sqlite3
    conn = sqlite3.connect("timesheets.db", check_same_thread=False)

c = conn.cursor()
# Ensure table exists
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
conn.commit()

# ==== Helper: Normalize names ====
def normalize_name(name: str) -> str:
    nfkd = unicodedata.normalize("NFKD", name)
    only_ascii = nfkd.encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z]", "", only_ascii).lower()

# ==== Load pay‑rate database (from path or uploaded file) ====
def load_rate_database(source):
    """
    `source` can be a filesystem path (str/Path) or a file‐like (uploaded).
    Returns: (custom_rates, normalized_rates, norm_to_raw)
    """
    custom_rates = {}
    normalized_rates = {}
    norm_to_raw = {}
    wb = load_workbook(source, data_only=True)
    for sheet in wb.sheetnames:
        # Read everything as DataFrame with no header to find where "Name" / "Pay Rate" live
        df_raw = pd.read_excel(source, sheet_name=sheet, header=None)
        header_row = None
        for idx, val in enumerate(df_raw.iloc[:,0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                second = df_raw.iat[idx, 1]
                if isinstance(second, str) and second.strip().lower() == "pay rate":
                    header_row = idx
                    break
        if header_row is None:
            continue
        df = pd.read_excel(source, sheet_name=sheet, header=header_row)
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

# ==== Sidebar: Upload pay_rates.xlsx or fallback ====
RATE_FILE_PATH = "pay_rates.xlsx"
rate_uploader = st.sidebar.file_uploader(
    "➕ Upload pay_rates.xlsx (optional)", type=["xlsx"]
)

try:
    if rate_uploader:
        custom_rates, normalized_rates, norm_to_raw = load_rate_database(rate_uploader)
    else:
        if not Path(RATE_FILE_PATH).exists():
            raise FileNotFoundError
        custom_rates, normalized_rates, norm_to_raw = load_rate_database(RATE_FILE_PATH)
except FileNotFoundError:
    st.sidebar.warning(
        "No pay_rates.xlsx found; defaulting to £15/hr for everyone. "
        "Upload one above to enable custom rates."
    )
    custom_rates, normalized_rates, norm_to_raw = {}, {}, {}

# ==== Lookup match or default ====
def lookup_match(name: str):
    norm = normalize_name(name)
    if norm in normalized_rates:
        return norm_to_raw[norm], normalized_rates[norm], 1.0
    return name, 15.0, 0.0

# ==== Placeholders for your existing parsing logic ====
def extract_from_docx(file) -> list[dict]:
    """Your DOCX‐parsing returning a list of dicts per person/day."""
    # … existing code …
    return []

def extract_from_pdf(file) -> list[dict]:
    """Your PDF‐parsing returning a list of dicts per person/day."""
    # … existing code …
    return []

def hhmm_to_hours(hhmm: str) -> float:
    try:
        h, m = hhmm.split(":")
        return int(h) + int(m) / 60.0
    except:
        return 0.0

def calculate_pay(name: str, daily_data: list[dict]) -> dict:
    """
    Summarize weekday/sat/sun hours for a given name
    and return a dict including date_range, hours, rate, etc.
    """
    # … existing code …
    return {}

# ==== Sidebar: Timesheet uploads ====
st.sidebar.header("Upload Timesheets")
st.sidebar.markdown("""
1. Upload **.docx** or **.pdf**  
2. Confirm name‑matches (expand Debug)  
3. Export Excel with formulas  
""")
uploaded_files = st.sidebar.file_uploader(
    "Choose timesheet files", accept_multiple_files=True
)

# ==== Main Tabs ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1) Upload & Review ----
with tabs[0]:
    st.header("📤 Upload & Review Timesheets")
    if uploaded_files:
        st.success(f"Processing {len(uploaded_files)} file(s)…")
        # … your parsing, match‐confirmation, debug expander, save to DB, export logic …
    else:
        st.info("Waiting for you to upload some .docx or .pdf timesheets…")

# ---- 2) History (date‐filtered + weekly summaries) ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Filter by upload date, then see a weekly summary:")

    # Date range picker (last 30 days default)
    today = date.today()
    start_date, end_date = st.date_input(
        "Select upload date range",
        value=(today - timedelta(days=30), today),
        min_value=date(2020, 1, 1),
        max_value=today
    )

    # Fetch all rows
    c.execute("""
        SELECT name, matched_as, ratio, client, site_address, department,
               weekday_hours, saturday_hours, sunday_hours, rate AS rate,
               date_range, extracted_on, source_file, upload_timestamp
        FROM timesheet_entries
        ORDER BY upload_timestamp DESC
    """)
    rows = c.fetchall()
    cols = [
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ]
    df = pd.DataFrame(rows, columns=cols)
    df["Upload Timestamp"] = pd.to_datetime(df["Upload Timestamp"]).dt.date

    # Apply date filter
    mask = (df["Upload Timestamp"] >= start_date) & (df["Upload Timestamp"] <= end_date)
    filtered = df.loc[mask].copy()

    if filtered.empty:
        st.info("No entries found in that date range.")
    else:
        # Compute pay per entry
        filtered["Pay"] = (
            filtered["Weekday Hours"]
            + filtered["Saturday Hours"]
            + filtered["Sunday Hours"]
        ) * filtered["Rate (£)"]

        # Group by week
        summary = (
            filtered
            .groupby("Date Range")
            .agg(
                Entries=("Name","count"),
                Weekday_Hours=("Weekday Hours","sum"),
                Sat_Hours=("Saturday Hours","sum"),
                Sun_Hours=("Sunday Hours","sum"),
                Total_Pay=("Pay","sum")
            )
            .reset_index()
            .sort_values("Date Range", ascending=False)
        )

        st.markdown("### 📅 Weekly Summary")
        st.dataframe(summary, use_container_width=True)

        with st.expander(f"Show raw entries ({len(filtered)})"):
            st.dataframe(filtered.drop(columns=["Pay"]), use_container_width=True)

# ---- 3) Dashboard ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.markdown("Aggregate stats & charts for all stored timesheets.")
    # … your existing dashboard code …

# ---- 4) Settings ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Reload custom rates via the sidebar uploader.  
    - Missing pay_rates.xlsx → everyone defaults to £15/hr.  
    - For feature requests, reach out to your dev team.
    """)

