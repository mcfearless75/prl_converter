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
    );
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
    );
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
    st.rerun()  # ← updated from experimental_rerun()

custom_rates, normalized_rates, normalized_keys, norm_to_raw = load_rate_database(RATE_FILE_PATH)

def lookup_match(name: str):
    norm = normalize_name(name)
    if not norm:
        return None, DEFAULT_RATE, 0.0
    if norm in normalized_rates:
        return norm_to_raw[norm], normalized_rates[norm], 1.0
    return None, DEFAULT_RATE, 0.0

def hhmm_to_float(hhmm: str) -> float:
    try:
        h, m = hhmm.strip().split(":")
        return int(h) + int(m) / 60.0
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
    overtime = max(0.0, wd - 50.0)
    pay_regular = (wd - overtime) * rate
    pay_overtime = overtime * rate * 1.5
    pay_sat = sat * rate * 1.5
    pay_sun = sun * rate * 1.75
    total_pay = pay_regular + pay_overtime + pay_sat + pay_sun
    return wd, sat, sun, rate, total_pay, matched_raw or "No match", ratio

def extract_timesheet_data(file) -> dict:
    # … your existing DOCX logic from the old app …
    # returns a dict matching the original structure :contentReference[oaicite:2]{index=2}
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # … your existing PDF logic from the old app …
    return []

# ====== Streamlit Tabs UI ======
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1. Upload & Review ----
with tabs[0]:
    st.sidebar.header("Upload Timesheets")
    st.sidebar.markdown(
        """
        1. Upload **.docx** or **.pdf** timesheet files.  
        2. Confirm name‐match (expand Debug).  
        3. Export Excel with formulas.
        """
    )
    uploaded_files = st.sidebar.file_uploader(
        "Select files", type=["docx", "pdf"], accept_multiple_files=True
    )
    if uploaded_files:
        all_rows = []
        progress = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            lower = file.name.lower()
            if lower.endswith(".docx"):
                rec = extract_timesheet_data(file)
                if not rec.get("Name"):
                    rec["Name"] = Path(file.name).stem.replace("_", " ").replace("-", " ").title()
                all_rows.append(rec)
            elif lower.endswith(".pdf"):
                pdf_list = extract_timesheet_data_pdf(file)
                for r in pdf_list:
                    if r.get("Name"):
                        all_rows.append(r)
            progress.progress((i + 1) / total_files)

        df = pd.DataFrame(all_rows)

        # 🔍 Debug table
        debug_df = (
            df[["Name", "Matched As", "Ratio", "Rate (£)", "Source File"]]
            .drop_duplicates()
            .reset_index(drop=True)
        )
        def highlight_low(val):
            return "background-color: #FFCCCC" if val == "No match" or (isinstance(val, float) and val < 1.0) else ""
        styled = debug_df.style.applymap(highlight_low, subset=["Ratio","Matched As"])
        with st.expander("Show Name-Match Debug Table"):
            st.dataframe(styled, use_container_width=True)

        # warnings & validation
        if (df["Matched As"] == "No match").any():
            st.warning("⚠️ Some names not matched, default rate applied.")
        problems = []
        for idx, row in df.iterrows():
            if row["Weekday Hours"] < 0 or row["Weekday Hours"] > 168:
                problems.append((idx, "Weekday hours out of range"))
            if row["Saturday Hours"] < 0 or row["Saturday Hours"] > 24:
                problems.append((idx, "Saturday hours out of range"))
            if row["Sunday Hours"] < 0 or row["Sunday Hours"] > 24:
                problems.append((idx, "Sunday hours out of range"))
        if problems:
            st.error("⚠️ Data issues:")
            for idx, reason in problems:
                st.write(f"- Row {idx+1} ({df.at[idx,'Name']}): {reason}")

        # Insert into DB
        if not problems:
            insert_q = (
                """INSERT INTO timesheet_entries
                   (name, matched_as, ratio, client, site_address, department,
                    weekday_hours, saturday_hours, sunday_hours, rate,
                    date_range, extracted_on, source_file, upload_timestamp)
                   VALUES ({})"""
            )
            # build placeholders per DB type
            for _, row in df.iterrows():
                params = (
                    row["Name"], row["Matched As"], row["Ratio"], row["Client"],
                    row["Site Address"], row["Department"], row["Weekday Hours"],
                    row["Saturday Hours"], row["Sunday Hours"], row["Rate (£)"],
                    row["Date Range"], row["Extracted On"], row["Source File"],
                    datetime.now()
                )
                if isinstance(conn, sqlite3.Connection):
                    ph = ",".join("?"*14)
                    c.execute(insert_q.format(ph), params)
                else:
                    ph = ",".join("%s"*14)
                    c.execute(insert_q.format(ph), params)
            conn.commit()
            st.success(f"✅ Inserted {len(df)} rows into history.")

        # Final table & summary
        st.markdown("### 📋 Final Timesheet Table")
        if "Calculated Pay (£)" not in df.columns:
            df["Calculated Pay (£)"] = df.apply(lambda row: calculate_pay(row["Name"], row.get("daily", []))[4], axis=1)
        st.dataframe(df, use_container_width=True)

        summary_df = (
            df.groupby("Matched As")[["Calculated Pay (£)", "Weekday Hours", "Saturday Hours", "Sunday Hours"]]
            .sum()
            .reset_index()
            .rename(columns={"Matched As":"Name"})
        )
        col1, col2 = st.columns([3,1])
        with col1:
            st.markdown("### 💰 Weekly Summary")
            st.dataframe(summary_df, use_container_width=True)
        with col2:
            total_hours = summary_df[["Weekday Hours","Saturday Hours","Sunday Hours"]].sum().sum()
            total_pay   = summary_df["Calculated Pay (£)"].sum()
            st.metric("Total Hours This Period", f"{total_hours:.2f}")
            st.metric("Total Pay This Period", f"£{total_pay:,.2f}")

        # Download Excel with formulas
        output = BytesIO()
        wb_out = Workbook()
        ws = wb_out.active; ws.title = "Timesheets"
        headers = [
            "Name","Matched As","Ratio","Client","Site Address","Department",
            "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
            "Regular Pay (£)","Overtime Pay (£)","Saturday Pay (£)","Sunday Pay (£)","Total Pay (£)",
            "Date Range","Extracted On","Source File"
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        for idx, row in df.iterrows():
            excel_row = idx + 2
            wd, sat, sun, rate = row["Weekday Hours"], row["Saturday Hours"], row["Sunday Hours"], row["Rate (£)"]
            reg_f  = f"=MIN(G{excel_row},50)*J{excel_row}"
            ot_f   = f"=MAX(G{excel_row}-50,0)*J{excel_row}*1.5"
            sat_f  = f"=H{excel_row}*J{excel_row}*1.5"
            sun_f  = f"=I{excel_row}*J{excel_row}*1.75"
            tot_f  = f"=K{excel_row}+L{excel_row}+M{excel_row}+N{excel_row}"
            ws.append([
                row["Name"], row["Matched As"], row["Ratio"], row["Client"], row["Site Address"], row["Department"],
                wd, sat, sun, rate,
                reg_f, ot_f, sat_f, sun_f, tot_f,
                row["Date Range"], row["Extracted On"], row["Source File"]
            ])
        for col in ws.columns:
            max_len = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2
        wb_out.save(output)
        st.download_button(
            "📥 Download Excel with Formulas",
            data=output.getvalue(),
            file_name="PRL_Timesheet_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---- 2. History (grouped by week) ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Past uploads grouped by week.")
    query = """
        SELECT name, matched_as, ratio, client, site_address, department,
               weekday_hours, saturday_hours, sunday_hours, rate,
               date_range, extracted_on, source_file, upload_timestamp
          FROM timesheet_entries
      ORDER BY upload_timestamp DESC
      LIMIT 1000
    """
    c.execute(query)
    rows = c.fetchall()
    history_df = pd.DataFrame(rows, columns=[
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ])
    history_df["Upload Timestamp"] = pd.to_datetime(history_df["Upload Timestamp"])
    history_df["week_start"] = (
        history_df["Upload Timestamp"]
        .dt.to_period("W-MON")
        .apply(lambda r: r.start_time.date())
    )
    for week, wk_df in history_df.groupby("week_start"):
        with st.expander(f"Week of {week.strftime('%Y-%m-%d')} ({len(wk_df)} entries)"):
            st.dataframe(wk_df.sort_values("Upload Timestamp").drop(columns=["week_start"]), use_container_width=True)

# ---- 3. Dashboard ----
with tabs[2]:
    st.header("📊 Dashboard")
    st.markdown("Aggregate stats for all stored timesheets.")
    c.execute("""
        SELECT matched_as, SUM(weekday_hours), SUM(saturday_hours), SUM(sunday_hours),
               SUM(weekday_hours*rate + saturday_hours*rate + sunday_hours*rate) AS total_pay
          FROM timesheet_entries
      GROUP BY matched_as
    """)
    dashboard_df = pd.DataFrame(c.fetchall(), columns=[
        "Name","Weekday Hours","Saturday Hours","Sunday Hours","Total Pay"
    ])
    if not dashboard_df.empty:
        st.bar_chart(dashboard_df.set_index("Name")[["Weekday Hours","Saturday Hours","Sunday Hours"]])
        st.markdown(f"**Total Pay (approx):** £{dashboard_df['Total Pay'].sum():,.2f}")
    else:
        st.info("No data available yet.")

# ---- 4. Settings ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Click **Reload Pay Rates** in the sidebar to re‑load your Excel.
    - Name matching is case‑ and accent‑insensitive.
    - Unmatched names default to £15/hr and are highlighted in red.
    """)
