import os
import io
import zipfile
from pathlib import Path
import streamlit as st
import pandas as pd
import docx
import pdfplumber
import re
import unicodedata
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
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
else:
    import sqlite3
    conn = sqlite3.connect("prl_timesheets.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        upload_timestamp TEXT
    )
    """)
conn.commit()

st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")

RATE_FILE_PATH = "pay details.xlsx"
DEFAULT_RATE = 15.0

def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

@st.cache_data
def load_rate_database(excel_path: str):
    wb = load_workbook(excel_path, read_only=True)
    custom_rates = {}
    normalized_rates = {}
    norm_to_raw = {}
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
    return custom_rates, normalized_rates, list(normalized_rates.keys()), norm_to_raw

if st.sidebar.button("🔄 Reload Pay Rates"):
    st.cache_data.clear()
    st.experimental_rerun()

custom_rates, normalized_rates, normalized_keys, norm_to_raw = load_rate_database(RATE_FILE_PATH)

def lookup_match(name: str):
    norm = normalize_name(name)
    if not norm:
        return None, DEFAULT_RATE, 0.0
    if norm in normalized_rates:
        raw = norm_to_raw[norm]
        return raw, normalized_rates[norm], 1.0
    return None, DEFAULT_RATE, 0.0

def hhmm_to_float(hhmm: str) -> float:
    try:
        h, m = hhmm.strip().split(":")
        return int(h) + int(m) / 60.0
    except:
        return 0.0

def calculate_pay(name: str, daily_data: list[dict]):
    matched_raw, rate, ratio = lookup_match(name)
    weekday_hours = sat_hours = sun_hours = 0.0
    for entry in daily_data:
        h = entry["hours"]
        wd = entry["weekday"]
        if wd == "Saturday":
            sat_hours += h
        elif wd == "Sunday":
            sun_hours += h
        else:
            weekday_hours += h
    overtime = max(0.0, weekday_hours - 50.0)
    pay_regular = (weekday_hours - overtime) * rate
    pay_overtime = overtime * rate * 1.5
    pay_sat = sat_hours * rate * 1.5
    pay_sun = sun_hours * rate * 1.75
    total_pay = pay_regular + pay_overtime + pay_sat + pay_sun
    return weekday_hours, sat_hours, sun_hours, rate, total_pay, matched_raw, ratio

def extract_timesheet_data(file) -> list[dict]:
    # ... your existing DOCX parsing logic here ...
    pass

def extract_timesheet_data_pdf(file) -> list[dict]:
    # ... your existing PDF parsing logic here ...
    pass

# ====== Streamlit Tabs UI ======
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1. Upload & Review ----
with tabs[0]:
    st.header("📥 Bulk‑Upload via ZIP")
    st.markdown("Upload a ZIP containing folders of `.pdf` and `.docx` timesheets.")

    uploaded_zip = st.file_uploader(
        "📦 Select your ZIP file",
        type=["zip"],
        accept_multiple_files=False
    )

    if uploaded_zip:
        try:
            zip_bytes = io.BytesIO(uploaded_zip.read())
            z = zipfile.ZipFile(zip_bytes)
        except zipfile.BadZipFile:
            st.error("That doesn’t look like a valid ZIP. Try re‑zipping and uploading again.")
        else:
            processed = 0
            for member in z.namelist():
                if member.lower().endswith((".pdf", ".docx")):
                    # read each file into BytesIO so your extractors can handle it
                    content = z.read(member)
                    file_obj = BytesIO(content)
                    file_obj.name = Path(member).name

                    # choose extractor
                    if member.lower().endswith(".pdf"):
                        records = extract_timesheet_data_pdf(file_obj)
                    else:
                        records = extract_timesheet_data(file_obj)

                    # --- your existing processing logic goes here ---
                    # e.g., loop over `records`, calculate pay,
                    # insert rows into DB, show a preview, etc.
                    #
                    # for rec in records:
                    #     weekday, sat, sun, rate, total, matched, ratio = calculate_pay(rec["name"], rec["daily"])
                    #     # insert into DB...
                    #     c.execute( ... )
                    # conn.commit()
                    #
                    processed += 1

            st.success(f"Processed {processed} files from the ZIP!")
    else:
        st.info("Upload a ZIP above to start processing timesheets.")

# ---- 2. History ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Displays all timesheet entries stored in the database.")

    query = (
        "SELECT name, matched_as, ratio, client, site_address, department, "
        "weekday_hours, saturday_hours, sunday_hours, rate, date_range, "
        "extracted_on, source_file, upload_timestamp "
        "FROM timesheet_entries "
        "ORDER BY upload_timestamp DESC LIMIT 1000"
    )
    c.execute(query)
    rows = c.fetchall()
    history_df = pd.DataFrame(
        rows,
        columns=[
            "Name", "Matched As", "Ratio", "Client", "Site Address", "Department",
            "Weekday Hours", "Saturday Hours", "Sunday Hours", "Rate (£)",
            "Date Range", "Extracted On", "Source File", "Upload Timestamp"
        ]
    )

    history_df["Upload Timestamp"] = pd.to_datetime(history_df["Upload Timestamp"])
    history_df["week_start"] = (
        history_df["Upload Timestamp"]
        .dt.to_period("W-MON")
        .apply(lambda r: r.start_time.date())
    )

    for week, wk_df in history_df.groupby("week_start"):
        with st.expander(f"Week of {week.strftime('%Y-%m-%d')} ({len(wk_df)} entries)"):
            st.dataframe(
                wk_df.sort_values("Upload Timestamp")
                     .drop(columns=["week_start"]),
                use_container_width=True
            )

# ---- 3. Dashboard ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.markdown("Aggregate stats for all stored timesheets.")
    # ... your existing dashboard code ...

# ---- 4. Settings ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - ZIP upload will recurse through folders and pick up any `.pdf` or `.docx`.
    - Click **Reload Pay Rates** in the sidebar to refresh from Excel.
    - Only exact matches (case & accent‑insensitive) pull custom rates.
    - Unmatched names default to £15/hr and get highlighted for review.
    """)
