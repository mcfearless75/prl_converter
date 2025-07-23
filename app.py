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

# ==== DB CONNECTION: Postgres on Render, SQLite locally ====
if "DATABASE_URL" in os.environ:
    import psycopg2
    from urllib.parse import urlparse
    url = urlparse(os.environ["DATABASE_URL"])
    conn = psycopg2.connect(
        dbname=url.path[1:], user=url.username,
        password=url.password, host=url.hostname, port=url.port
    )
    c = conn.cursor()
    placeholder = "%s"
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
        id SERIAL PRIMARY KEY,
        name TEXT, matched_as TEXT, ratio REAL,
        client TEXT, site_address TEXT, department TEXT,
        weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
        default_rate REAL, default_pay REAL,
        paul_rate REAL, paul_pay REAL,
        date_range TEXT, extracted_on TEXT, source_file TEXT,
        upload_timestamp TIMESTAMP
    );
    """)
else:
    import sqlite3
    conn = sqlite3.connect("prl_timesheets.db", check_same_thread=False)
    c = conn.cursor()
    placeholder = "?"
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, matched_as TEXT, ratio REAL,
        client TEXT, site_address TEXT, department TEXT,
        weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
        default_rate REAL, default_pay REAL,
        paul_rate REAL, paul_pay REAL,
        date_range TEXT, extracted_on TEXT, source_file TEXT,
        upload_timestamp TEXT
    );
    """)
conn.commit()

st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")

# ==== RATE FILES & DEFAULTS ====
DEFAULT_RATE_FILE = "pay details.xlsx"
PAUL_RATE_FILE   = "PAUL RATES.xlsx"
DEFAULT_RATE     = 15.0

# ==== HELPERS: NORMALIZE NAMES ====
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

# ==== LOAD RATES ====
@st.cache_data
def load_rate_database(path: str):
    if not os.path.exists(path):
        return {}, {}
    wb = load_workbook(path, read_only=True)
    rates = {}
    norm2raw = {}
    for sheet in wb.sheetnames:
        df0 = pd.read_excel(path, sheet_name=sheet, header=None)
        header_row = None
        for i, val in enumerate(df0.iloc[:,0]):
            if isinstance(val, str) and val.strip().lower()=="name" \
               and isinstance(df0.iat[i,1], str) and df0.iat[i,1].strip().lower()=="pay rate":
                header_row = i
                break
        if header_row is None:
            continue
        df = pd.read_excel(path, sheet_name=sheet, header=header_row)
        if not {"Name","Pay Rate"}.issubset(df.columns):
            continue
        df = df[["Name","Pay Rate"]].copy()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df = df.dropna(subset=["Name","Pay Rate"])
        for _, row in df.iterrows():
            raw = str(row["Name"]).strip()
            rate = float(row["Pay Rate"])
            norm = normalize_name(raw)
            rates[norm] = rate
            norm2raw[norm] = raw
    return rates, norm2raw

default_rates, default_norm2raw = load_rate_database(DEFAULT_RATE_FILE)
paul_rates,   paul_norm2raw   = load_rate_database(PAUL_RATE_FILE)

def lookup_rate(name: str, rates: dict, norm2raw: dict):
    norm = normalize_name(name)
    if norm in rates:
        return norm2raw[norm], rates[norm]
    return None, DEFAULT_RATE

def compute_pay(hours: float, rate: float, overtime_threshold=50.0):
    overtime = max(0.0, hours - overtime_threshold)
    regular = hours - overtime
    return regular*rate + overtime*rate*1.5

# ==== ORIGINAL EXTRACTORS (unchanged) ====
def hhmm_to_float(hhmm: str) -> float:
    try:
        h, m = hhmm.split(":")
        return int(h) + int(m)/60.0
    except:
        return 0.0

def calculate_pay(name: str, daily_data: list[dict]):
    # unused now
    ...

def extract_timesheet_data(file) -> dict:
    # original DOCX logic from app.py22.txt
    # (omitted for brevity, paste in your full implementation)
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # original PDF logic from app.py22.txt
    # (omitted for brevity, paste in your full implementation)
    return []

# ==== SIDEBAR: RELOAD RATE FILES ====
if st.sidebar.button("🔄 Reload Rate Files"):
    st.cache_data.clear()
    st.experimental_rerun()

# ==== UI TABS ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ UPLOAD & REVIEW ----
with tabs[0]:
    st.header("📥 Upload Timesheets (.docx/.pdf/.zip)")
    uploaded = st.file_uploader(
        "Select files or a ZIP",
        type=["docx","pdf","zip"], accept_multiple_files=True
    )
    if uploaded:
        rows = []
        prog = st.progress(0)
        for i, f in enumerate(uploaded):
            name = f.name.lower()
            sources = []
            if name.endswith(".zip"):
                try:
                    zp = zipfile.ZipFile(io.BytesIO(f.read()))
                except zipfile.BadZipFile:
                    st.error(f"❌ Bad ZIP: {f.name}")
                    continue
                for member in zp.namelist():
                    if member.lower().endswith((".docx",".pdf")):
                        bio = BytesIO(zp.read(member)); bio.name = Path(member).name
                        sources.append(bio)
            else:
                sources.append(f)
            for src in sources:
                if src.name.lower().endswith(".docx"):
                    recs = [extract_timesheet_data(src)]
                else:
                    recs = extract_timesheet_data_pdf(src)
                for rec in recs:
                    # compute hours + pays
                    wd = rec["Weekday Hours"]
                    sat = rec["Saturday Hours"]
                    sun = rec["Sunday Hours"]
                    # default
                    _, dr = lookup_rate(rec["Name"], default_rates, default_norm2raw)
                    default_pay = compute_pay(wd, dr) + sat*dr*1.5 + sun*dr*1.75
                    # paul
                    _, pr = lookup_rate(rec["Name"], paul_rates, paul_norm2raw)
                    paul_pay = compute_pay(wd, pr) + sat*pr*1.5 + sun*pr*1.75
                    rows.append({
                        **rec,
                        "Default Rate (£)": dr,
                        "Default Pay (£)": default_pay,
                        "Paul Rate (£)": pr,
                        "Paul Pay (£)": paul_pay
                    })
            prog.progress((i+1)/len(uploaded))

        df = pd.DataFrame(rows)
        st.success(f"✅ Processed {len(rows)} records")

        # show table
        st.dataframe(df, use_container_width=True)

        # export button
        buf = BytesIO()
        wb = Workbook()
        ws = wb.active; ws.title="Report"
        ws.append(list(df.columns))
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        for row in df.itertuples(index=False):
            ws.append(list(row))
        wb.save(buf)
        st.download_button(
            "📥 Download Excel Report",
            data=buf.getvalue(),
            file_name="timesheet_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---- 2️⃣ HISTORY ----
with tabs[1]:
    st.header("🗃️ Upload History (weekly)")
    c.execute("""
      SELECT name, matched_as, ratio, client, site_address, department,
             weekday_hours, saturday_hours, sunday_hours,
             default_rate, default_pay, paul_rate, paul_pay,
             date_range, extracted_on, source_file, upload_timestamp
        FROM timesheet_entries
       ORDER BY upload_timestamp DESC LIMIT 1000
    """)
    hist = pd.DataFrame(c.fetchall(), columns=[
      "Name","Matched As","Ratio","Client","Site Address","Department",
      "Weekday Hours","Saturday Hours","Sunday Hours",
      "Default Rate (£)","Default Pay (£)","Paul Rate (£)","Paul Pay (£)",
      "Date Range","Extracted On","Source File","Uploaded At"
    ])
    hist["Uploaded At"] = pd.to_datetime(hist["Uploaded At"])
    hist["Week Start"] = hist["Uploaded At"].dt.to_period("W-MON").apply(lambda r: r.start_time.date())
    for wk, grp in hist.groupby("Week Start"):
        with st.expander(f"Week of {wk} ({len(grp)} recs)"):
            st.dataframe(grp.drop(columns=["Week Start"]), use_container_width=True)

# ---- 3️⃣ DASHBOARD ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.markdown("Compare Default vs Paul total pay")
    if not hist.empty:
        agg = hist.groupby("Name")[["Default Pay (£)","Paul Pay (£)"]].sum()
        st.bar_chart(agg)
    else:
        st.info("No data yet.")

# ---- 4️⃣ SETTINGS ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Click **Reload Rate Files** after uploading new rate sheets.
    - Upload timesheets in `.docx`, `.pdf` or `.zip`.
    - Table and Excel export now include both Default and Paul pays.
    - History tab shows your saved uploads grouped by week.
    """)
