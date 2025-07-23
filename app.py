import os
import io
import zipfile
import sqlite3
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

# ==== DB Connection: Postgres if available, else SQLite ====
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
      id SERIAL PRIMARY KEY, name TEXT, matched_as TEXT, ratio REAL,
      client TEXT, site_address TEXT, department TEXT,
      weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
      rate REAL, date_range TEXT, extracted_on TEXT,
      source_file TEXT, upload_timestamp TIMESTAMP
    );
    """)
else:
    conn = sqlite3.connect("prl_timesheets.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, matched_as TEXT,
      ratio REAL, client TEXT, site_address TEXT, department TEXT,
      weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
      rate REAL, date_range TEXT, extracted_on TEXT,
      source_file TEXT, upload_timestamp TEXT
    );
    """)
conn.commit()

# ==== Page config & Rate file ====
st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")
RATE_FILE_PATH = "pay details.xlsx"
DEFAULT_RATE = 15.0

# ==== Sidebar: Reload pay rates ====
if st.sidebar.button("🔄 Reload Pay Rates"):
    st.cache_data.clear()
    st.rerun()

# ==== Name normalization & rate DB loader ====
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

@st.cache_data
def load_rate_database(excel_path: str):
    wb = load_workbook(excel_path, read_only=True)
    custom_rates = {}
    normalized_rates = {}
    norm_to_raw = {}
    for sheet in wb.sheetnames:
        df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        header_row = None
        for idx, val in enumerate(df_raw.iloc[:,0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                second = df_raw.iat[idx,1]
                if isinstance(second, str) and second.strip().lower() == "pay rate":
                    header_row = idx
                    break
        if header_row is None:
            continue
        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
        if not {"Name","Pay Rate"}.issubset(df.columns):
            continue
        df = df[["Name","Pay Rate"]].copy()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df = df.dropna(subset=["Name","Pay Rate"])
        for _,row in df.iterrows():
            raw = str(row["Name"]).strip()
            rate = float(row["Pay Rate"])
            custom_rates[raw] = rate
            norm = normalize_name(raw)
            normalized_rates[norm] = rate
            norm_to_raw[norm] = raw
    return custom_rates, normalized_rates, list(normalized_rates.keys()), norm_to_raw

custom_rates, normalized_rates, normalized_keys, norm_to_raw = load_rate_database(RATE_FILE_PATH)

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
    matched_raw, rate, ratio = lookup_match(name)
    wd = sat = sun = 0.0
    for entry in daily_data:
        h = entry["hours"]
        day = entry["weekday"]
        if day == "Saturday":
            sat += h
        elif day == "Sunday":
            sun += h
        else:
            wd += h
    overtime = max(0, wd - 50)
    reg_pay = (wd - overtime) * rate
    ot_pay  = overtime * rate * 1.5
    sat_pay = sat * rate * 1.5
    sun_pay = sun * rate * 1.75
    total   = reg_pay + ot_pay + sat_pay + sun_pay
    return wd, sat, sun, rate, total, matched_raw or "No match", ratio

# ====== Stubs for your extractors ======
def extract_timesheet_data(file) -> dict:
    # …your DOCX logic here…
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # …your PDF logic here…
    return []

# ====== Streamlit Tabs UI ======
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ Upload & Review ----
with tabs[0]:
    st.sidebar.header("Upload Timesheets")
    st.sidebar.markdown(
        "1. Upload **.docx**, **.pdf**, or a **.zip** of them.\n"
        "2. Expand **Debug** to inspect name-matching.\n"
        "3. Export Excel with formulas and save to History."
    )
    uploaded_files = st.sidebar.file_uploader(
        "Select files or ZIP", type=["docx","pdf","zip"], accept_multiple_files=True
    )

    if uploaded_files:
        all_rows = []
        progress = st.progress(0)
        total = len(uploaded_files)
        for i, f in enumerate(uploaded_files):
            name = f.name.lower()
            # ZIP → recurse
            if name.endswith(".zip"):
                try:
                    zb = zipfile.ZipFile(io.BytesIO(f.read()))
                except zipfile.BadZipFile:
                    st.error(f"Bad ZIP: {f.name}")
                    continue
                for member in zb.namelist():
                    if member.lower().endswith((".docx",".pdf")):
                        data = zb.read(member)
                        bio = BytesIO(data); bio.name = Path(member).name
                        if member.lower().endswith(".docx"):
                            rec = extract_timesheet_data(bio)
                            if not rec.get("Name"):
                                rec["Name"] = Path(bio.name).stem.title()
                            all_rows.append(rec)
                        else:
                            for rec in extract_timesheet_data_pdf(bio):
                                all_rows.append(rec)
            # Single DOCX/PDF
            elif name.endswith(".docx"):
                rec = extract_timesheet_data(f)
                if not rec.get("Name"):
                    rec["Name"] = Path(f.name).stem.title()
                all_rows.append(rec)
            elif name.endswith(".pdf"):
                for rec in extract_timesheet_data_pdf(f):
                    all_rows.append(rec)

            progress.progress((i+1)/total)

        df = pd.DataFrame(all_rows)
        st.success(f"Processed {len(all_rows)} records from {total} uploads")

        # 🔍 Debug table
        if not df.empty:
            debug_df = df[["Name","Matched As","Ratio","Rate (£)","Source File"]].drop_duplicates()
            def hl(val):
                if val=="No match" or (isinstance(val,(int,float)) and val<1.0):
                    return "background-color: #FFCCCC"
                return ""
            styled = debug_df.style.applymap(hl, subset=["Matched As","Ratio"])
            with st.expander("Show Name‑Match Debug Table"):
                st.dataframe(styled, use_container_width=True)

            # warnings
            if (df["Matched As"]=="No match").any():
                st.warning("⚠️ Some names not matched; default rate used.")

            # validation
            problems = []
            for idx,row in df.iterrows():
                if row.get("Weekday Hours",0) <0 or row.get("Weekday Hours",0)>168:
                    problems.append(f"Row {idx+1}: Weekday out of range")
                if row.get("Saturday Hours",0)<0 or row.get("Saturday Hours",0)>24:
                    problems.append(f"Row {idx+1}: Saturday out of range")
                if row.get("Sunday Hours",0)<0 or row.get("Sunday Hours",0)>24:
                    problems.append(f"Row {idx+1}: Sunday out of range")
            if problems:
                st.error("⚠️ Data validation issues:")
                for p in problems: st.write("- "+p)

            # Insert into DB
            if not problems:
                for idx,row in df.iterrows():
                    params = (
                        row["Name"], row["Matched As"], row["Ratio"], row["Client"],
                        row["Site Address"], row["Department"], row["Weekday Hours"],
                        row["Saturday Hours"], row["Sunday Hours"], row["Rate (£)"],
                        row["Date Range"], row["Extracted On"], row["Source File"],
                        datetime.now()
                    )
                    if isinstance(conn, sqlite3.Connection):
                        ph = ",".join("?"*14)
                        c.execute(f"INSERT INTO timesheet_entries VALUES (NULL,{ph})", params)
                    else:
                        ph = ",".join("%s"*14)
                        c.execute(f"INSERT INTO timesheet_entries VALUES (DEFAULT,{ph})", params)
                conn.commit()
                st.success(f"✅ Inserted {len(df)} rows into history.")

            # Final table & summary
            if "Calculated Pay (£)" not in df.columns:
                df["Calculated Pay (£)"] = df.apply(
                    lambda r: calculate_pay(r["Name"], r.get("daily",[]))[4], axis=1
                )
            st.markdown("### 📋 Final Timesheet Table")
            st.dataframe(df, use_container_width=True)

            summary_df = (
                df.groupby("Matched As")[["Calculated Pay (£)","Weekday Hours","Saturday Hours","Sunday Hours"]]
                  .sum().reset_index().rename(columns={"Matched As":"Name"})
            )
            col1,col2 = st.columns([3,1])
            with col1:
                st.markdown("### 💰 Summary by Name")
                st.dataframe(summary_df, use_container_width=True)
            with col2:
                tot_hrs = summary_df[["Weekday Hours","Saturday Hours","Sunday Hours"]].sum().sum()
                tot_pay = summary_df["Calculated Pay (£)"].sum()
                st.metric("Total Hours", f"{tot_hrs:.2f}")
                st.metric("Total Pay", f"£{tot_pay:,.2f}")

            # Download Excel with formulas
            output = BytesIO()
            wb_out = Workbook()
            ws = wb_out.active; ws.title = "Timesheets"
            headers = [
                "Name","Matched As","Ratio","Client","Site Address","Department",
                "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
                "Regular Pay","Overtime Pay","Saturday Pay","Sunday Pay","Total Pay",
                "Date Range","Extracted On","Source File"
            ]
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(fill_type="solid", start_color="D9D9D9", end_color="D9D9D9")
            for idx,row in df.iterrows():
                er = idx+2
                regf = f"=MIN(G{er},50)*J{er}"
                otf  = f"=MAX(G{er}-50,0)*J{er}*1.5"
                sf   = f"=H{er}*J{er}*1.5"
                sunf = f"=I{er}*J{er}*1.75"
                tf   = f"=K{er}+L{er}+M{er}+N{er}"
                ws.append([
                    row["Name"], row["Matched As"], row["Ratio"], row["Client"],
                    row["Site Address"], row["Department"],
                    row["Weekday Hours"], row["Saturday Hours"], row["Sunday Hours"],
                    row["Rate (£)",],
                    regf, otf, sf, sunf, tf,
                    row["Date Range"], row["Extracted On"], row["Source File"]
                ])
            wb_out.save(output)
            st.download_button(
                "📥 Download Excel with Formulas",
                data=output.getvalue(),
                file_name="PRL_Timesheet_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("ℹ️ No records found—check your uploads or extractor logic.")

# ---- 2️⃣ History (grouped by week) ----
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
    history_df = pd.DataFrame(rows, columns=[
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ])
    history_df["Upload Timestamp"] = pd.to_datetime(history_df["Upload Timestamp"])
    history_df["week_start"] = history_df["Upload Timestamp"].dt.to_period("W-MON").apply(lambda r: r.start_time.date())
    for wk,grp in history_df.groupby("week_start"):
        with st.expander(f"Week of {wk} ({len(grp)} entries)"):
            st.dataframe(grp.sort_values("Upload Timestamp").drop(columns="week_start"), use_container_width=True)

# ---- 3️⃣ Dashboard ----
with tabs[2]:
    st.header("📊 Dashboard")
    c.execute("""
      SELECT matched_as,
             SUM(weekday_hours), SUM(saturday_hours), SUM(sunday_hours),
             SUM(weekday_hours*rate + saturday_hours*rate + sunday_hours*rate) as total_pay
        FROM timesheet_entries GROUP BY matched_as
    """)
    drows = c.fetchall()
    dashboard_df = pd.DataFrame(drows, columns=[
        "Name","Weekday Hours","Saturday Hours","Sunday Hours","Total Pay"
    ])
    if not dashboard_df.empty:
        st.bar_chart(dashboard_df.set_index("Name")[["Weekday Hours","Saturday Hours","Sunday Hours"]])
        st.markdown(f"**Total Pay (approx):** £{dashboard_df['Total Pay'].sum():,.2f}")
    else:
        st.info("No data available yet.")

# ---- 4️⃣ Settings ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Click **Reload Pay Rates** to re‑load your Excel.
    - Upload `.zip` as well as `.docx`/`.pdf`.
    - Name matching is case‑ and accent‑insensitive.
    - Unmatched names default to £15/hr.
    """)
