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
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# ==== CONFIG & PAGE SETUP ====
st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")
RATE_FILE_PATH = "pay details.xlsx"
DEFAULT_RATE = 15.0

# ==== SIDEBAR: PAY‑RATES EXCEL UPLOADER ====
st.sidebar.title("⚙️ PRL Timesheet Converter")
excel_uploader = st.sidebar.file_uploader(
    "📥 Upload Pay‑Rates Excel",
    type=["xlsx"],
    help="Select a new pay‑details.xlsx to replace current rates."
)
if excel_uploader:
    with open(RATE_FILE_PATH, "wb") as f:
        f.write(excel_uploader.getbuffer())
    st.sidebar.success("✅ Pay rates file updated!")
    st.cache_data.clear()
    st.rerun()

# ==== DATABASE CONNECTION ====
if "DATABASE_URL" in os.environ:
    import psycopg2
    from urllib.parse import urlparse
    url = urlparse(os.environ["DATABASE_URL"])
    conn = psycopg2.connect(
        dbname=url.path[1:], user=url.username,
        password=url.password, host=url.hostname, port=url.port
    )
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
        id SERIAL PRIMARY KEY,
        name TEXT, matched_as TEXT, ratio REAL, client TEXT,
        site_address TEXT, department TEXT,
        weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
        rate REAL, date_range TEXT, extracted_on TEXT,
        source_file TEXT, upload_timestamp TIMESTAMP
    );
    """)
else:
    import sqlite3
    conn = sqlite3.connect("prl_timesheets.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, matched_as TEXT, ratio REAL, client TEXT,
        site_address TEXT, department TEXT,
        weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
        rate REAL, date_range TEXT, extracted_on TEXT,
        source_file TEXT, upload_timestamp TEXT
    );
    """)
conn.commit()

# ==== UTILITIES: NAME NORMALIZATION & RATE LOADING ====
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

@st.cache_data
def load_rate_database(excel_path: str):
    wb = load_workbook(excel_path, read_only=True)
    normalized_rates: dict[str, float] = {}
    norm_to_raw: dict[str, str] = {}

    for sheet in wb.sheetnames:
        # read raw to find header row
        df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        header_row = None
        for idx, val in enumerate(df_raw.iloc[:, 0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                second = df_raw.iat[idx, 1]
                if isinstance(second, str) and second.strip().lower() == "pay rate":
                    header_row = idx
                    break
        if header_row is None:
            continue

        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
        if not {"Name", "Pay Rate"}.issubset(df.columns):
            continue

        df = df[["Name", "Pay Rate"]].copy()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df = df.dropna(subset=["Name", "Pay Rate"])

        for _, row in df.iterrows():
            raw_name = str(row["Name"]).strip()
            rate = float(row["Pay Rate"])
            norm = normalize_name(raw_name)
            normalized_rates[norm] = rate
            norm_to_raw[norm] = raw_name

    return normalized_rates, norm_to_raw

# Load rates once per run (cached)
normalized_rates, norm_to_raw = load_rate_database(RATE_FILE_PATH)

def lookup_match(name: str):
    norm = normalize_name(name)
    if not norm or norm not in normalized_rates:
        return None, DEFAULT_RATE, 0.0
    return norm_to_raw[norm], normalized_rates[norm], 1.0

def hhmm_to_float(hhmm: str) -> float:
    try:
        h, m = hhmm.split(":")
        return int(h) + int(m)/60
    except:
        return 0.0

def calculate_pay(name: str, daily_data: list[dict]):
    matched, rate, ratio = lookup_match(name)
    wd = sat = sun = 0.0
    for e in daily_data:
        h = e["hours"]
        d = e["weekday"]
        if d == "Saturday":
            sat += h
        elif d == "Sunday":
            sun += h
        else:
            wd += h
    overtime = max(0, wd - 50)
    total = (wd - overtime) * rate + overtime * rate*1.5 + sat * rate*1.5 + sun * rate*1.75
    return wd, sat, sun, rate, total, matched, ratio

def extract_timesheet_data(file) -> list[dict]:
    # your DOCX parsing logic goes here
    # return a list of dicts like:
    # [{"name": ..., "daily": [...], "client": ..., "site_address": ..., "department": ...}, ...]
    return []

def extract_timesheet_data_pdf(file) -> list[dict]:
    # your PDF parsing logic goes here (same return structure)
    return []

# ==== STREAMLIT UI: TABS ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ UPLOAD & REVIEW ----
with tabs[0]:
    st.header("📥 Bulk‑Upload via ZIP")
    st.write("Upload a ZIP containing subfolders of `.pdf` & `.docx` timesheets.")
    uploaded_zip = st.file_uploader("Select ZIP file", type="zip")

    if uploaded_zip:
        try:
            z = zipfile.ZipFile(io.BytesIO(uploaded_zip.read()))
        except zipfile.BadZipFile:
            st.error("Invalid ZIP. Please re‑zip and retry.")
        else:
            all_records: list[dict] = []
            file_count = 0

            for member in z.namelist():
                if member.lower().endswith((".pdf", ".docx")):
                    raw = z.read(member)
                    bio = BytesIO(raw)
                    bio.name = Path(member).name
                    file_count += 1

                    # extract
                    if member.lower().endswith(".pdf"):
                        recs = extract_timesheet_data_pdf(bio)
                    else:
                        recs = extract_timesheet_data(bio)

                    # tag & collect
                    for rec in recs:
                        rec["source_file"] = bio.name
                        rec["upload_timestamp"] = datetime.now()
                        all_records.append(rec)

            st.success(f"Processed {file_count} files & {len(all_records)} records!")

            if all_records:
                df = pd.DataFrame(all_records)
                st.subheader("🔍 Preview of extracted entries")
                st.dataframe(df, use_container_width=True)

                if st.button("💾 Save all to database"):
                    # parameter style differs between SQLite & Postgres; unify via named style
                    for rec in all_records:
                        if isinstance(conn, sqlite3.Connection):
                            placeholders = ",".join("?" * len(rec))
                            columns = ",".join(rec.keys())
                            c.execute(
                                f"INSERT INTO timesheet_entries ({columns}) VALUES ({placeholders})",
                                tuple(rec.values())
                            )
                        else:
                            # psycopg2 named style
                            cols = ", ".join(rec.keys())
                            vals = ", ".join(f"%({k})s" for k in rec.keys())
                            c.execute(
                                f"INSERT INTO timesheet_entries ({cols}) VALUES ({vals})",
                                rec
                            )
                    conn.commit()
                    st.success(f"Saved {len(all_records)} records to the DB!")
            else:
                st.info("🤔 No timesheet entries were found in your files.")

# ---- 2️⃣ HISTORY ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    query = """
        SELECT name, matched_as, ratio, client, site_address, department,
               weekday_hours, saturday_hours, sunday_hours, rate,
               date_range, extracted_on, source_file, upload_timestamp
          FROM timesheet_entries
      ORDER BY upload_timestamp DESC LIMIT 1000
    """
    c.execute(query)
    rows = c.fetchall()
    df = pd.DataFrame(rows, columns=[
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ])

    df["Upload Timestamp"] = pd.to_datetime(df["Upload Timestamp"])
    df["week_start"] = df["Upload Timestamp"].dt.to_period("W-MON").apply(lambda r: r.start_time.date())

    for wk, sub in df.groupby("week_start"):
        with st.expander(f"Week of {wk} ({len(sub)} entries)"):
            st.dataframe(
                sub.sort_values("Upload Timestamp")
                   .drop(columns="week_start"),
                use_container_width=True
            )

# ---- 3️⃣ DASHBOARD ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.write("Aggregate stats for all stored timesheets.")
    # …your existing dashboard charts/code…

# ---- 4️⃣ SETTINGS ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - ZIP upload will recurse through subfolders for `.pdf` & `.docx`.
    - Use the sidebar “Upload Pay‑Rates Excel” to replace `pay details.xlsx`.
    - Unmatched names default to £15/hr.
    - After upload & preview, click “Save all” to persist entries.
    """)
