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
    nfkd = unicodedata.normalize("NFKD", name)
    only_ascii = nfkd.encode("ASCII", "ignore").decode("utf-8")
    return re.sub(r"[^a-zA-Z]", "", only_ascii).lower()

def load_rate_database(source):
    custom, normed, to_raw = {}, {}, {}
    wb = load_workbook(source, data_only=True)
    for sheet in wb.sheetnames:
        df0 = pd.read_excel(source, sheet_name=sheet, header=None)
        header_row = None
        for i, v in enumerate(df0.iloc[:, 0]):
            if (
                isinstance(v, str)
                and v.strip().lower() == "name"
                and isinstance(df0.iat[i, 1], str)
                and df0.iat[i, 1].strip().lower() == "pay rate"
            ):
                header_row = i
                break
        if header_row is None:
            continue
        df = pd.read_excel(source, sheet_name=sheet, header=header_row)
        if "Name" not in df or "Pay Rate" not in df:
            continue
        df = df[["Name", "Pay Rate"]].dropna()
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        for _, r in df.iterrows():
            raw = str(r["Name"]).strip()
            rate = float(r["Pay Rate"])
            custom[raw] = rate
            n = normalize_name(raw)
            normed[n] = rate
            to_raw[n] = raw
    return custom, normed, to_raw

# ==== Rate‐Loader (multiple sheets) ====
RATE_FILE = "pay_rates.xlsx"
uploads = st.sidebar.file_uploader(
    "➕ Upload one or more pay‑rate XLSX files",
    type=["xlsx"], accept_multiple_files=True
)

custom_rates, normalized_rates, norm_to_raw = {}, {}, {}
def _merge(src):
    c, n, t = load_rate_database(src)
    custom_rates.update(c)
    normalized_rates.update(n)
    norm_to_raw.update(t)

if uploads:
    for f in uploads:
        _merge(f)
    st.sidebar.success(f"Merged {len(uploads)} rate sheet(s).")
elif Path(RATE_FILE).exists():
    try:
        _merge(RATE_FILE)
        st.sidebar.info(f"Loaded local `{RATE_FILE}`.")
    except Exception as e:
        st.sidebar.error(f"Error loading `{RATE_FILE}`: {e}")
if not normalized_rates:
    st.sidebar.warning(
        "No pay‑rate data found; everyone will default to £15/hr. "
        "Upload XLSX(s) above to enable custom rates."
    )

def lookup_match(name: str):
    n = normalize_name(name)
    if n in normalized_rates:
        return norm_to_raw[n], normalized_rates[n], 1.0
    return name, 15.0, 0.0

# ==== Extraction Logic ====
def extract_from_docx(file) -> list[dict]:
    doc = docx.Document(file)
    if not doc.tables:
        return []
    tbl = doc.tables[0]
    header = tbl.rows[0].cells[0].text.split("\n")
    header = [h.strip() for h in header if h.strip()]
    client = name = site = None
    for i, line in enumerate(header):
        low = line.lower()
        if low.startswith("client"):
            client = line.split(None, 1)[1].strip()
            if i + 1 < len(header):
                name = header[i + 1].strip()
        if low.startswith("site address"):
            parts = line.split("\t", 1)
            site = parts[1].strip() if len(parts) > 1 else None
    header_row = None
    for idx, row in enumerate(tbl.rows):
        if row.cells[0].text.strip().lower() == "date":
            header_row = idx
            break
    if header_row is None:
        return []
    date_re = re.compile(r"\d{2}\.\d{2}\.\d{4}")
    weekday = saturday = sunday = 0.0
    dates_list = []
    for row in tbl.rows[header_row + 1 :]:
        dt = row.cells[0].text.strip()
        if not date_re.match(dt):
            continue
        dates_list.append(dt)
        day = row.cells[1].text.strip().lower()
        try:
            hrs = float(row.cells[4].text.strip())
        except:
            hrs = 0.0
        if day == "saturday":
            saturday += hrs
        elif day == "sunday":
            sunday += hrs
        else:
            weekday += hrs
    if not dates_list:
        return []
    start, end = min(dates_list), max(dates_list)
    drange = f"{start}–{end}"
    matched, rate, ratio = lookup_match(name or "")
    return [
        {
            "name": name or "",
            "matched_as": matched,
            "ratio": ratio,
            "client": client or "",
            "site_address": site or "",
            "department": "",
            "weekday_hours": weekday,
            "saturday_hours": saturday,
            "sunday_hours": sunday,
            "rate": rate,
            "date_range": drange,
            "extracted_on": datetime.now().isoformat(),
        }
    ]

def extract_from_pdf(file) -> list[dict]:
    return []

# ==== Sidebar: Timesheet uploader ====
st.sidebar.header("Upload Timesheets")
st.sidebar.markdown(
    """
1. Upload **.docx**, **.pdf** or **.zip**  
2. Confirm name‑matches (expand Debug)  
3. Export Excel with formulas  
"""
)
uploaded = st.sidebar.file_uploader(
    "Choose timesheet file(s)", accept_multiple_files=True
)

# ==== Main Tabs ====
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ---- 1) Upload & Review ----
with tabs[0]:
    st.header("📤 Upload & Review Timesheets")

    if not uploaded:
        st.info("Waiting for you to upload .docx, .pdf or .zip…")
    else:
        total = len(uploaded)
        progress = st.progress(0)
        summaries = []

        with st.spinner(f"Processing {total} file(s)…"):
            for i, uf in enumerate(uploaded):
                st.write(f"➡️ Handling **{uf.name}** ({i+1}/{total})")
                lower = uf.name.lower()
                if lower.endswith(".zip"):
                    try:
                        z = zipfile.ZipFile(uf)
                        members = [
                            f
                            for f in z.namelist()
                            if f.lower().endswith((".docx", ".pdf"))
                        ]
                        if not members:
                            st.warning(f"{uf.name} had no .docx/.pdf inside.")
                        for m in members:
                            st.write(f" • Extracting `{m}`")
                            data = z.read(m)
                            buf = BytesIO(data)
                            buf.name = m
                            recs = (
                                extract_from_docx(buf)
                                if m.lower().endswith(".docx")
                                else extract_from_pdf(buf)
                            )
                            st.write(
                                f"   → Got {len(recs)} summary record(s)"
                            )
                            summaries += recs
                    except zipfile.BadZipFile:
                        st.error(f"{uf.name} is not a valid ZIP.")
                elif lower.endswith(".docx"):
                    st.write(" • Parsing DOCX…")
                    recs = extract_from_docx(uf)
                    st.write(
                        f"   → Got {len(recs)} summary record(s}"
                    )
                    summaries += recs
                elif lower.endswith(".pdf"):
                    st.write(" • Parsing PDF…")
                    recs = extract_from_pdf(uf)
                    st.write(
                        f"   → Got {len(recs)} summary record(s)"
                    )
                    summaries += recs
                else:
                    st.warning(f"Unsupported file: {uf.name}")

                progress.progress((i + 1) / total)

        if not summaries:
            st.error("No valid timesheet records were extracted.")
        else:
            st.success(f"✅ Extracted {len(summaries)} weekly summary record(s).")
            df = pd.DataFrame(summaries)
            with st.expander("🔍 Debug: Raw summaries"):
                st.dataframe(df, use_container_width=True)

            # --- EXCEL EXPORT SNIPPET ----
            towrite = BytesIO()
            with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Summaries")
            towrite.seek(0)

            st.download_button(
                label="📥 Download Timesheet Summary (Excel)",
                data=towrite,
                file_name=f"timesheet_summary_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            # --------------------------------

            # --- (Optional) Insert into DB and your original export logic ---
            # for rec in summaries:
            #     c.execute(
            #         "INSERT INTO timesheet_entries (...) VALUES (…) ",
            #         (…)
            #     )
            # conn.commit()

# ---- 2) History (date‑filter + weekly summary) ----
with tabs[1]:
    st.header("🗃️ Timesheet Upload History")
    st.markdown("Filter by upload date, then see a weekly summary:")

    today = date.today()
    start_date, end_date = st.date_input(
        "Select upload date range",
        value=(today - timedelta(days=30), today),
        min_value=date(2020, 1, 1),
        max_value=today
    )

    c.execute(
        """
        SELECT name, matched_as, ratio, client, site_address, department,
               weekday_hours, saturday_hours, sunday_hours, rate AS rate,
               date_range, extracted_on, source_file, upload_timestamp
        FROM timesheet_entries
        ORDER BY upload_timestamp DESC
        """
    )
    rows = c.fetchall()
    cols = [
        "Name","Matched As","Ratio","Client","Site Address","Department",
        "Weekday Hours","Saturday Hours","Sunday Hours","Rate (£)",
        "Date Range","Extracted On","Source File","Upload Timestamp"
    ]
    hist = pd.DataFrame(rows, columns=cols)
    hist["Upload Timestamp"] = pd.to_datetime(hist["Upload Timestamp"]).dt.date

    mask = (
        (hist["Upload Timestamp"] >= start_date)
        & (hist["Upload Timestamp"] <= end_date)
    )
    filt = hist.loc[mask]

    if filt.empty:
        st.info("No entries found in that date range.")
    else:
        filt["Pay"] = (
            filt["Weekday Hours"]
            + filt["Saturday Hours"]
            + filt["Sunday Hours"]
        ) * filt["Rate (£)"]

        summary = (
            filt.groupby("Date Range")
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

        with st.expander(f"Show raw entries ({len(filt)})"):
            st.dataframe(filt.drop(columns=["Pay"]), use_container_width=True)

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
