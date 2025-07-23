import os
import io
import zipfile
import sqlite3
import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ==== DB CONNECTION & SCHEMA ==== 
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

# ==== MIGRATION: add new columns if missing (SQLite) ====
if isinstance(conn, sqlite3.Connection):
    cols = [r[1] for r in c.execute("PRAGMA table_info(timesheet_entries)").fetchall()]
    for col in ("default_rate","default_pay","paul_rate","paul_pay"):
        if col not in cols:
            c.execute(f"ALTER TABLE timesheet_entries ADD COLUMN {col} REAL")
    conn.commit()

# ==== PAGE + RATE FILES ==== 
st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")
DEFAULT_RATE_FILE = "pay details.xlsx"
PAUL_RATE_FILE   = "PAUL RATES.xlsx"
DEFAULT_RATE     = 15.0

# ==== HELPERS ==== 
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

@st.cache_data
def load_rate_database(path: str):
    if not os.path.exists(path):
        return {}, {}
    wb = load_workbook(path, read_only=True)
    rates, norm2raw = {}, {}
    for sheet in wb.sheetnames:
        df0 = pd.read_excel(path, sheet_name=sheet, header=None)
        hdr = None
        for i, v in enumerate(df0.iloc[:,0]):
            if isinstance(v,str) and v.strip().lower()=="name" \
               and isinstance(df0.iat[i,1],str) and df0.iat[i,1].strip().lower()=="pay rate":
                hdr = i; break
        if hdr is None:
            continue
        df = pd.read_excel(path, sheet_name=sheet, header=hdr)
        if not {"Name","Pay Rate"}.issubset(df.columns): 
            continue
        df = df[["Name","Pay Rate"]].dropna(subset=["Name","Pay Rate"])
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
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

def compute_pay(hours, rate, threshold=50.0):
    overtime = max(0.0, hours - threshold)
    regular = hours - overtime
    return regular*rate + overtime*rate*1.5

# ==== STUB EXTRACTORS ==== 
def extract_timesheet_data(file) -> dict:
    # your DOCX logic here
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # your PDF logic here
    return []

# ==== SIDEBAR: RELOAD RATE FILES ==== 
if st.sidebar.button("🔄 Reload Rate Files"):
    st.cache_data.clear()
    st.rerun()

# ==== UI TABS ==== 
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ UPLOAD & REVIEW ----
with tabs[0]:
    st.header("📥 Upload Timesheets `.docx` / `.pdf` / `.zip`")
    uploads = st.file_uploader(
        "Select files or a ZIP", type=["docx","pdf","zip"], accept_multiple_files=True
    )

    if uploads:
        records = []
        prog = st.progress(0)
        total = len(uploads)
        for i, f in enumerate(uploads):
            nm = f.name.lower()
            buffers = []
            if nm.endswith(".zip"):
                try:
                    zp = zipfile.ZipFile(io.BytesIO(f.read()))
                except zipfile.BadZipFile:
                    st.error(f"❌ Bad ZIP: {f.name}")
                    continue
                for member in zp.namelist():
                    if member.lower().endswith((".docx",".pdf")):
                        buf = BytesIO(zp.read(member)); buf.name = Path(member).name
                        buffers.append(buf)
            else:
                buffers.append(f)

            for buf in buffers:
                if buf.name.lower().endswith(".docx"):
                    recs = [extract_timesheet_data(buf)]
                else:
                    recs = extract_timesheet_data_pdf(buf)
                for rec in recs:
                    # DEFENSIVE ACCESS OF HOURS
                    wd  = rec.get("Weekday Hours", rec.get("weekday_hours", 0))
                    sat = rec.get("Saturday Hours", rec.get("saturday_hours", 0))
                    sun = rec.get("Sunday Hours", rec.get("sunday_hours", 0))

                    # default pay
                    _, dr = lookup_rate(rec.get("Name",""), default_rates, default_norm2raw)
                    dp = compute_pay(wd, dr) + sat*dr*1.5 + sun*dr*1.75

                    # paul pay
                    _, pr = lookup_rate(rec.get("Name",""), paul_rates, paul_norm2raw)
                    pp = compute_pay(wd, pr) + sat*pr*1.5 + sun*pr*1.75

                    records.append({
                        **rec,
                        "Default Rate (£)": dr,
                        "Default Pay (£)": dp,
                        "Paul Rate (£)": pr,
                        "Paul Pay (£)": pp
                    })

            prog.progress((i+1)/total)

        df = pd.DataFrame(records)
        st.success(f"✅ Processed {len(records)} records")

        # SHOW & EXPORT
        st.dataframe(df, use_container_width=True)

        buf = BytesIO()
        wb = Workbook()
        ws = wb.active; ws.title = "Comparison"
        ws.append(list(df.columns))
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(fill_type="solid", start_color="D9D9D9", end_color="D9D9D9")
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
    st.header("🗃️ Upload History (Weekly)")
    c.execute(f"""
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
    if not hist.empty:
        agg = hist.groupby("Name")[["Default Pay (£)","Paul Pay (£)"]].sum()
        st.bar_chart(agg)
    else:
        st.info("No history yet.")

# ---- 4️⃣ SETTINGS & INFO ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Upload new rate sheets via the Sidebar and click **Reload Rate Files**.
    - Import `.docx`, `.pdf`, or `.zip` files in Upload & Review.
    - Table and Excel export compare Default vs Paul pay.
    - History groups past uploads by week.
    """)
