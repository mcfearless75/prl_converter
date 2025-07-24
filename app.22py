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

# ==== DB CONNECTION & SCHEMA CREATION ====
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
      paul_rate REAL, paul_ot_rate REAL, paul_pay REAL,
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
      paul_rate REAL, paul_ot_rate REAL, paul_pay REAL,
      date_range TEXT, extracted_on TEXT, source_file TEXT,
      upload_timestamp TEXT
    );
    """)
conn.commit()

# ==== SCHEMA MIGRATION (SQLite only) ====
if isinstance(conn, sqlite3.Connection):
    existing = [r[1] for r in c.execute("PRAGMA table_info(timesheet_entries)").fetchall()]
    for col in ("default_rate","default_pay","paul_rate","paul_ot_rate","paul_pay"):
        if col not in existing:
            c.execute(f"ALTER TABLE timesheet_entries ADD COLUMN {col} REAL")
    conn.commit()

# ==== PAGE CONFIG & RATE FILE PATHS ====
st.set_page_config(page_title="PRL Timesheet Converter", page_icon="📑", layout="wide")
DEFAULT_RATE_FILE = "pay details.xlsx"
PAUL_RATE_FILE   = "PAUL RATES.xlsx"
DEFAULT_RATE     = 15.0

# ==== HELPERS: NAME NORMALIZATION ====
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

# ==== RATE LOADERS ====
@st.cache_data
def load_default_rates(path: str):
    rates, norm2raw = {}, {}
    if not os.path.exists(path):
        return rates, norm2raw
    wb = load_workbook(path, read_only=True)
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

@st.cache_data
def load_paul_rates(path: str):
    base_rates, ot_rates, norm2raw = {}, {}, {}
    if not os.path.exists(path):
        return base_rates, ot_rates, norm2raw
    wb = load_workbook(path, read_only=True)
    for sheet in wb.sheetnames:
        df0 = pd.read_excel(path, sheet_name=sheet, header=None)
        hdr = None
        for i, v in enumerate(df0.iloc[:,0]):
            if isinstance(v,str) and v.strip().lower()=="name":
                hdr = i; break
        if hdr is None:
            continue
        df = pd.read_excel(path, sheet_name=sheet, header=hdr)
        # Expect columns "Name","Pay Rate","OT Rate"
        if not {"Name","Pay Rate","OT Rate"}.issubset(df.columns):
            continue
        df = df[["Name","Pay Rate","OT Rate"]].dropna(subset=["Name"])
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df["OT Rate"]  = pd.to_numeric(df["OT Rate"], errors="coerce")
        for _, row in df.iterrows():
            raw = str(row["Name"]).strip()
            br  = float(row["Pay Rate"])
            orate = float(row["OT Rate"])
            norm = normalize_name(raw)
            base_rates[norm] = br
            ot_rates[norm]   = orate
            norm2raw[norm]   = raw
    return base_rates, ot_rates, norm2raw

default_rates, default_norm2raw = load_default_rates(DEFAULT_RATE_FILE)
paul_base,   paul_ot_rates, paul_norm2raw = load_paul_rates(PAUL_RATE_FILE)

# ==== PAY CALCULATION ====
def compute_pay(
    name: str,
    daily: list[dict],
    base_rates: dict,
    norm2raw: dict,
    ot_rates: dict | None = None
):
    norm = normalize_name(name)
    rate = base_rates.get(norm, DEFAULT_RATE)
    wd = sat = sun = 0.0
    for e in daily:
        h = e.get("hours",0.0)
        d = e.get("weekday","").lower()
        if d.startswith("sat"):
            sat += h
        elif d.startswith("sun"):
            sun += h
        else:
            wd += h
    overtime = max(0.0, wd - 50.0)
    regular = wd - overtime
    if ot_rates and norm in ot_rates:
        orate = ot_rates[norm]
        pay = regular*rate + overtime*orate + sat*orate + sun*orate
    else:
        pay = regular*rate + overtime*rate*1.5 + sat*rate*1.5 + sun*rate*1.75
    return wd, sat, sun, rate, pay

# ==== STUB EXTRACTORS ====
def extract_timesheet_data(file) -> dict:
    # … your DOCX logic …
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # … your PDF logic …
    return []

# ==== SIDEBAR: RELOAD RATE SHEETS ====
if st.sidebar.button("🔄 Reload Rate Files"):
    st.cache_data.clear()
    st.rerun()

# ==== UI TABS ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ UPLOAD & REVIEW ----
with tabs[0]:
    st.header("📥 Upload Timesheets (.docx/.pdf/.zip)")
    uploads = st.file_uploader(
        "Select files or a ZIP",
        type=["docx","pdf","zip"], accept_multiple_files=True
    )
    if uploads:
        records = []
        prog = st.progress(0)
        total = len(uploads)
        for i, f in enumerate(uploads):
            nm = f.name.lower()
            sources = []
            if nm.endswith(".zip"):
                try:
                    zp = zipfile.ZipFile(io.BytesIO(f.read()))
                except zipfile.BadZipFile:
                    st.error(f"❌ Bad ZIP: {f.name}")
                    continue
                for member in zp.namelist():
                    if member.lower().endswith((".docx",".pdf")):
                        buf = BytesIO(zp.read(member)); buf.name = Path(member).name
                        sources.append(buf)
            else:
                sources.append(f)

            for src in sources:
                if src.name.lower().endswith(".docx"):
                    recs = [extract_timesheet_data(src)]
                else:
                    recs = extract_timesheet_data_pdf(src)
                for rec in recs:
                    wd, sat, sun, dr, dp = compute_pay(
                        rec.get("Name",""), rec.get("daily",[]),
                        default_rates, default_norm2raw
                    )
                    _, pr = normalize_name(rec.get("Name","")), DEFAULT_RATE  # fallback
                    wd2, sat2, sun2, pr, pp = compute_pay(
                        rec.get("Name",""), rec.get("daily",[]),
                        paul_base, paul_norm2raw, paul_ot_rates
                    )
                    records.append({
                        **rec,
                        "Default Rate (£)": dr,
                        "Default Pay (£)": dp,
                        "Paul Rate (£)": pr,
                        "Paul OT Rate (£)": paul_ot_rates.get(normalize_name(rec.get("Name","")), pr*1.5),
                        "Paul Pay (£)": pp
                    })
            prog.progress((i+1)/total)

        df = pd.DataFrame(records)
        st.success(f"✅ Processed {len(records)} records")

        st.dataframe(df, use_container_width=True)

        buf = BytesIO()
        wb = Workbook(); ws = wb.active; ws.title = "Comparison"
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
             default_rate, default_pay, paul_rate, paul_ot_rate, paul_pay,
             date_range, extracted_on, source_file, upload_timestamp
        FROM timesheet_entries
      ORDER BY upload_timestamp DESC LIMIT 1000
    """)
    hist = pd.DataFrame(c.fetchall(), columns=[
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours",
        "Default Rate (£)","Default Pay (£)","Paul Rate (£)","Paul OT Rate (£)","Paul Pay (£)",
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
    - Upload new rate sheets via the sidebar and click Reload.
    - Timesheet uploads support .docx, .pdf, and .zip.
    - Comparison table shows both default and Paul pays.
    - Export to Excel and review History by week.
    """)
