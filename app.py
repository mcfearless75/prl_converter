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

# ==== PAGE CONFIG & RATE FILE ====
st.set_page_config(page_title="PRL Timesheet Portal", page_icon="📑", layout="wide")
RATE_FILE_PATH = "pay details.xlsx"
DEFAULT_RATE = 15.0

# ==== SIDEBAR: UPLOAD PAY‑RATES EXCEL ====
st.sidebar.title("⚙️ PRL Timesheet Converter")
excel_uploader = st.sidebar.file_uploader(
    "📥 Upload Pay‑Rates Excel",
    type=["xlsx"],
    help="Drop in a fresh pay details.xlsx to update rates."
)
if excel_uploader:
    with open(RATE_FILE_PATH, "wb") as f:
        f.write(excel_uploader.getbuffer())
    st.sidebar.success("✅ Pay‑rates file updated!")
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
      name TEXT, matched_as TEXT, ratio REAL,
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
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT, matched_as TEXT, ratio REAL,
      client TEXT, site_address TEXT, department TEXT,
      weekday_hours REAL, saturday_hours REAL, sunday_hours REAL,
      rate REAL, date_range TEXT, extracted_on TEXT,
      source_file TEXT, upload_timestamp TEXT
    );
    """)
conn.commit()

# ==== UTILITIES: NORMALIZE & LOAD RATES ====
def normalize_name(s: str) -> str:
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    return re.sub(r"\s+", " ", s)

@st.cache_data
def load_rate_database(excel_path: str):
    wb = load_workbook(excel_path, read_only=True)
    normalized_rates = {}
    norm_to_raw = {}
    for sheet in wb.sheetnames:
        df0 = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        header_row = None
        for i, val in enumerate(df0.iloc[:, 0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                if (
                    isinstance(df0.iat[i, 1], str)
                    and df0.iat[i, 1].strip().lower() == "pay rate"
                ):
                    header_row = i
                    break
        if header_row is None:
            continue
        df = pd.read_excel(excel_path, sheet_name=sheet, header=header_row)
        if not {"Name","Pay Rate"}.issubset(df.columns):
            continue
        df = df[["Name","Pay Rate"]].copy()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        df = df.dropna(subset=["Name","Pay Rate"])
        for _, row in df.iterrows():
            raw = str(row["Name"]).strip()
            rate = float(row["Pay Rate"])
            norm = normalize_name(raw)
            normalized_rates[norm] = rate
            norm_to_raw[norm] = raw
    return normalized_rates, norm_to_raw

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
    matched_raw, rate, ratio = lookup_match(name)
    wd = sat = sun = 0.0
    for e in daily_data:
        hrs = e.get("hours", 0)
        day = e.get("weekday", "")
        if day == "Saturday":
            sat += hrs
        elif day == "Sunday":
            sun += hrs
        else:
            wd += hrs
    overtime = max(0, wd - 50)
    reg = (wd - overtime) * rate
    ot  = overtime * rate * 1.5
    sp  = sat * rate * 1.5
    sup = sun * rate * 1.75
    total = reg + ot + sp + sup
    return wd, sat, sun, rate, total, matched_raw or "No match", ratio

# ==== STUBS FOR YOUR EXTRACTORS ====
def extract_timesheet_data(file) -> dict:
    # TODO: implement DOCX parsing
    return {}

def extract_timesheet_data_pdf(file) -> list[dict]:
    # TODO: implement PDF parsing
    return []

# ==== STREAMLIT UI: TABS ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1️⃣ UPLOAD & REVIEW ----
with tabs[0]:
    st.header("📥 Upload Timesheet Files or ZIP")
    st.write("Upload `.docx`, `.pdf`, or a `.zip` containing them.")
    uploaded = st.file_uploader(
        "Select files or ZIP",
        type=["docx","pdf","zip"],
        accept_multiple_files=True
    )

    if uploaded:
        all_records = []
        prog = st.progress(0)
        for idx, f in enumerate(uploaded):
            nm = f.name.lower()
            # handle ZIP
            if nm.endswith(".zip"):
                try:
                    zp = zipfile.ZipFile(io.BytesIO(f.read()))
                except zipfile.BadZipFile:
                    st.error(f"❌ Bad ZIP: {f.name}")
                    continue
                for member in zp.namelist():
                    if member.lower().endswith((".docx",".pdf")):
                        data = zp.read(member)
                        bio = BytesIO(data); bio.name = Path(member).name
                        if member.lower().endswith(".docx"):
                            rec = extract_timesheet_data(bio)
                            if rec: all_records.append(rec)
                        else:
                            for rec in extract_timesheet_data_pdf(bio):
                                all_records.append(rec)
            # handle DOCX
            elif nm.endswith(".docx"):
                rec = extract_timesheet_data(f)
                if rec: all_records.append(rec)
            # handle PDF
            elif nm.endswith(".pdf"):
                for rec in extract_timesheet_data_pdf(f):
                    all_records.append(rec)

            prog.progress((idx+1)/len(uploaded))

        st.success(f"✅ Processed {len(uploaded)} uploads → {len(all_records)} records")

        if all_records:
            df = pd.DataFrame(all_records)

            # Defensive Debug Table
            debug_cols = ["Name","Matched As","Ratio","Rate (£)","Source File"]
            cols = [c for c in debug_cols if c in df.columns]
            if cols:
                dbg = df[cols].drop_duplicates()
                def hl(val):
                    if val=="No match" or (isinstance(val,(int,float)) and val<1.0):
                        return "background-color: #FFCCCC"
                    return ""
                styled = dbg.style.applymap(hl, subset=[c for c in cols if c in ("Matched As","Ratio")])
                with st.expander("Show Name‑Match Debug Table"):
                    st.dataframe(styled, use_container_width=True)
            else:
                st.info(f"No debug columns present. Found: {', '.join(df.columns)}")

            # Warnings
            if "Matched As" in df.columns and (df["Matched As"]=="No match").any():
                st.warning("⚠️ Some names not matched; default rate used.")

            # Validation
            problems = []
            for i, row in df.iterrows():
                for hcol, mx in [("Weekday Hours",168),("Saturday Hours",24),("Sunday Hours",24)]:
                    if hcol in row and (row[hcol]<0 or row[hcol]>mx):
                        problems.append(f"Row {i+1}: {hcol} out of range")
            if problems:
                st.error("⚠️ Data validation issues:")
                for p in problems: st.write(f"- {p}")

            # Insert into DB
            if not problems:
                for row in all_records:
                    wd, sat, sun, rate, total, ma, ratio = calculate_pay(row.get("name",""), row.get("daily",[]))
                    params = (
                        row.get("name",""), ma, ratio, row.get("client",""),
                        row.get("site_address",""), row.get("department",""),
                        wd, sat, sun, rate,
                        row.get("date_range",""), row.get("extracted_on",""),
                        row.get("source_file",""), datetime.now()
                    )
                    if isinstance(conn, sqlite3.Connection):
                        ph = ",".join("?"*14)
                        c.execute(f"""
                            INSERT INTO timesheet_entries
                            (name, matched_as, ratio, client, site_address, department,
                             weekday_hours, saturday_hours, sunday_hours, rate,
                             date_range, extracted_on, source_file, upload_timestamp)
                            VALUES ({ph})
                        """, params)
                    else:
                        ph = ",".join("%s"*14)
                        c.execute(f"""
                            INSERT INTO timesheet_entries
                            (name, matched_as, ratio, client, site_address, department,
                             weekday_hours, saturday_hours, sunday_hours, rate,
                             date_range, extracted_on, source_file, upload_timestamp)
                            VALUES ({ph})
                        """, params)
                conn.commit()
                st.success(f"✅ Saved {len(all_records)} records to history")

            # Final Preview
            if "Calculated Pay (£)" not in df.columns:
                df["Calculated Pay (£)"] = df.apply(
                    lambda r: calculate_pay(r.get("name",""), r.get("daily",[]))[4],
                    axis=1
                )
            st.markdown("### 📋 Final Timesheet Table")
            st.dataframe(df, use_container_width=True)

            # Summary Metrics
            summary = df.groupby("Matched As")[["Calculated Pay (£)","Weekday Hours","Saturday Hours","Sunday Hours"]].sum().reset_index().rename(columns={"Matched As":"Name"})
            c1, c2 = st.columns([3,1])
            with c1:
                st.markdown("### 💰 Summary by Name")
                st.dataframe(summary, use_container_width=True)
            with c2:
                total_hrs = summary[["Weekday Hours","Saturday Hours","Sunday Hours"]].sum().sum()
                total_pay = summary["Calculated Pay (£)"].sum()
                st.metric("Total Hours", f"{total_hrs:.2f}")
                st.metric("Total Pay", f"£{total_pay:,.2f}")

            # Excel Export with Formulas
            out = BytesIO()
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
            for i, row in df.iterrows():
                r = i+2
                regf = f"=MIN(G{r},50)*J{r}"
                otf  = f"=MAX(G{r}-50,0)*J{r}*1.5"
                sf   = f"=H{r}*J{r}*1.5"
                supf = f"=I{r}*J{r}*1.75"
                tf   = f"=K{r}+L{r}+M{r}+N{r}"
                ws.append([
                    row.get("name",""), row.get("Matched As",""), row.get("Ratio",""),
                    row.get("client",""), row.get("site_address",""), row.get("department",""),
                    row.get("Weekday Hours",0), row.get("Saturday Hours",0), row.get("Sunday Hours",0),
                    row.get("Rate (£)",0), regf, otf, sf, supf, tf,
                    row.get("date_range",""), row.get("extracted_on",""), row.get("source_file","")
                ])
            wb_out.save(out)
            st.download_button(
                "📥 Download Excel with Formulas",
                data=out.getvalue(),
                file_name="PRL_Timesheet_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.info("ℹ️ No records extracted – check your extractor logic.")

# ---- 2️⃣ HISTORY (GROUPED BY WEEK) ----
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
    hist_df = pd.DataFrame(rows, columns=[
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ])
    hist_df["Upload Timestamp"] = pd.to_datetime(hist_df["Upload Timestamp"])
    hist_df["week_start"] = hist_df["Upload Timestamp"].dt.to_period("W-MON").apply(lambda r: r.start_time.date())
    for wk, grp in hist_df.groupby("week_start"):
        with st.expander(f"Week of {wk} ({len(grp)} entries)"):
            st.dataframe(grp.sort_values("Upload Timestamp").drop(columns=["week_start"]), use_container_width=True)

# ---- 3️⃣ DASHBOARD ----
with tabs[2]:
    st.header("📊 Dashboard")
    c.execute("""
      SELECT matched_as,
             SUM(weekday_hours), SUM(saturday_hours), SUM(sunday_hours),
             SUM(weekday_hours*rate + saturday_hours*rate + sunday_hours*rate) AS total_pay
        FROM timesheet_entries
       GROUP BY matched_as
    """)
    drows = c.fetchall()
    dash_df = pd.DataFrame(drows, columns=[
        "Name","Weekday Hours","Saturday Hours","Sunday Hours","Total Pay"
    ])
    if not dash_df.empty:
        st.bar_chart(dash_df.set_index("Name")[["Weekday Hours","Saturday Hours","Sunday Hours"]])
        st.markdown(f"**Total Pay (approx):** £{dash_df['Total Pay'].sum():,.2f}")
    else:
        st.info("No data available yet.")

# ---- 4️⃣ SETTINGS & INFO ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Drop in a new Pay‑Rates Excel in the sidebar to update rates.
    - Upload `.docx`, `.pdf`, or a `.zip` of them in the Upload & Review tab.
    - Expand Debug to inspect name‑matching.
    - Export an Excel with live formulas and save to History.
    """)
