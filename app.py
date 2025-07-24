import os
import re
import unicodedata
import zipfile
from pathlib import Path
from io import BytesIO
from datetime import datetime, date, timedelta

import streamlit as st
import pandas as pd
import docx
import pdfplumber
from openpyxl import load_workbook

# ==== DB Connection (Postgres on Render, SQLite locally) ====
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

# ==== Helpers ====
def normalize_name(name: str) -> str:
    """Strip accents & non‑letters, lowercase."""
    nfkd = unicodedata.normalize("NFKD", name)
    only_ascii = nfkd.encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z]", "", only_ascii).lower()

def load_rate_database(source):
    """
    Load a single XLSX (path or file‑like),
    return (custom_rates, normalized_rates, norm_to_raw).
    """
    custom_rates, normalized_rates, norm_to_raw = {}, {}, {}
    wb = load_workbook(source, data_only=True)
    for sheet in wb.sheetnames:
        df_raw = pd.read_excel(source, sheet_name=sheet, header=None)
        header_row = None
        for idx, val in enumerate(df_raw.iloc[:, 0]):
            if isinstance(val, str) and val.strip().lower() == "name":
                if isinstance(df_raw.iat[idx, 1], str) and df_raw.iat[idx, 1].strip().lower() == "pay rate":
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

def lookup_match(name: str):
    """
    Return (matched_name, rate, confidence_ratio).
    Defaults to 15.0 if no custom rate found.
    """
    norm = normalize_name(name)
    if norm in normalized_rates:
        return norm_to_raw[norm], normalized_rates[norm], 1.0
    return name, 15.0, 0.0

def hhmm_to_hours(hhmm: str) -> float:
    try:
        h, m = hhmm.split(":")
        return int(h) + int(m)/60.0
    except:
        return 0.0

# Placeholders for your existing parsing logic:
def extract_from_docx(file) -> list[dict]:
    """Your DOCX parsing; returns list of daily records."""
    # … your code …
    return []

def extract_from_pdf(file) -> list[dict]:
    """Your PDF parsing; returns list of daily records."""
    # … your code …
    return []

def calculate_pay(name: str, daily_data: list[dict]) -> dict:
    """Aggregate weekday/sat/sun and produce row for DB."""
    # … your code …
    return {}

# ==== Sidebar: Multi‑sheet Rate‑Loader ====
RATE_FILE_PATH = "pay_rates.xlsx"
rate_uploaders = st.sidebar.file_uploader(
    "➕ Upload one or more pay‑rate XLSX files",
    type=["xlsx"],
    accept_multiple_files=True
)

custom_rates = {}
normalized_rates = {}
norm_to_raw = {}

def merge_rate_file(source):
    cr, nr, nt = load_rate_database(source)
    custom_rates.update(cr)
    normalized_rates.update(nr)
    norm_to_raw.update(nt)

# 1) Merge any uploaded rate sheets
if rate_uploaders:
    for up in rate_uploaders:
        merge_rate_file(up)
    st.sidebar.success(f"Merged {len(rate_uploaders)} rate sheet(s).")
# 2) Else try local fallback
elif Path(RATE_FILE_PATH).exists():
    try:
        merge_rate_file(RATE_FILE_PATH)
        st.sidebar.info(f"Loaded local `{RATE_FILE_PATH}`.")
    except Exception as e:
        st.sidebar.error(f"Error loading `{RATE_FILE_PATH}`: {e}")
# 3) Warn if still empty
if not normalized_rates:
    st.sidebar.warning(
        "No pay‑rate data found; defaulting to £15/hr for everyone. "
        "Upload XLSX(s) above to enable custom rates."
    )

# ==== Sidebar: Timesheet Upload UI ====
st.sidebar.header("Upload Timesheets")
st.sidebar.markdown("""
1. Upload **.docx**, **.pdf** or **.zip**  
2. Confirm name‑matches (expand Debug)  
3. Export Excel with formulas  
""")
uploaded_files = st.sidebar.file_uploader(
    "Choose timesheet file(s)", accept_multiple_files=True
)

# ==== Main Tabs ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1) Upload & Review ----
with tabs[0]:
    st.header("📤 Upload & Review Timesheets")

    if not uploaded_files:
        st.info("Waiting for you to upload .docx, .pdf or .zip of timesheets…")
    else:
        total = len(uploaded_files)
        progress = st.progress(0)
        all_records = []

        with st.spinner(f"Processing {total} file(s)…"):
            for idx, uf in enumerate(uploaded_files):
                st.write(f"➡️ Handling **{uf.name}** ({idx+1}/{total})")
                lower = uf.name.lower()

                # ZIP support
                if lower.endswith(".zip"):
                    try:
                        z = zipfile.ZipFile(uf)
                        members = [f for f in z.namelist() if f.lower().endswith((".docx", ".pdf"))]
                        if not members:
                            st.warning(f"❗ {uf.name} contains no .docx/.pdf files.")
                        for member in members:
                            st.write(f" • Extracting `{member}`")
                            data = z.read(member)
                            buff = BytesIO(data)
                            buff.name = member
                            if member.lower().endswith(".docx"):
                                recs = extract_from_docx(buff)
                            else:
                                recs = extract_from_pdf(buff)
                            st.write(f"   → Got {len(recs)} records")
                            all_records.extend(recs)
                    except zipfile.BadZipFile:
                        st.error(f"❌ {uf.name} is not a valid ZIP.")
                # DOCX
                elif lower.endswith(".docx"):
                    st.write(" • Parsing DOCX…")
                    recs = extract_from_docx(uf)
                    st.write(f"   → Got {len(recs)} records")
                    all_records.extend(recs)
                # PDF
                elif lower.endswith(".pdf"):
                    st.write(" • Parsing PDF…")
                    recs = extract_from_pdf(uf)
                    st.write(f"   → Got {len(recs)} records")
                    all_records.extend(recs)
                else:
                    st.warning(f"Unsupported file type: {uf.name}")

                progress.progress((idx + 1) / total)

        # Final result
        if all_records:
            st.success(f"✅ Finished! Total records extracted: {len(all_records)}")
            df = pd.DataFrame(all_records)
            with st.expander("🔍 Debug: Raw extraction DataFrame"):
                st.dataframe(df, use_container_width=True)
            # … your calculate_pay(), DB insert, and Excel‑export logic …
        else:
            st.error("No valid timesheet records were extracted.")

# ---- 2) History (Date‑filter + Weekly Summary) ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Filter by upload date, then see a weekly summary:")

    today = date.today()
    start_date, end_date = st.date_input(
        "Select upload date range",
        value=(today - timedelta(days=30), today),
        min_value=date(2020,1,1),
        max_value=today
    )

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

    mask = (df["Upload Timestamp"] >= start_date) & (df["Upload Timestamp"] <= end_date)
    filtered = df.loc[mask].copy()

    if filtered.empty:
        st.info("No entries found in that date range.")
    else:
        filtered["Pay"] = (
            filtered["Weekday Hours"]
            + filtered["Saturday Hours"]
            + filtered["Sunday Hours"]
        ) * filtered["Rate (£)"]

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

# ---- 4) Settings & Info ----
with tabs[3]:
    st.header("⚙️ Settings & Info")
    st.markdown("""
    - Drag‑n‑drop any number of rate spreadsheets; later uploads overwrite earlier rates.  
    - Missing rate data → everyone defaults to £15/hr.  
    - For feature requests or help, ping your dev team!
    """)
