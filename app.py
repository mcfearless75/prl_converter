import streamlit as st
import pandas as pd
import sqlite3
import docx
import pdfplumber
import re
import unicodedata
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from difflib import SequenceMatcher

# ──────────────────────────────────────────────────────────────────────────────
# 0) PAGE CONFIGURATION
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PRL Timesheet Portal",
    page_icon="📑",
    layout="wide"
)

# ──────────────────────────────────────────────────────────────────────────────
# 1) DATABASE SETUP (SQLite) FOR PERSISTENT STORAGE & HISTORY
# ──────────────────────────────────────────────────────────────────────────────
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

# ──────────────────────────────────────────────────────────────────────────────
# 2) GLOBAL SETTINGS & RATE LOADING
# ──────────────────────────────────────────────────────────────────────────────
RATE_FILE_PATH = "pay details.xlsx"
DEFAULT_RATE = 15.0
SIMILARITY_THRESHOLD = 0.60

def normalize_name(s: str) -> str:
    """
    Lowercase, strip accents, remove non-alphanumeric characters, collapse spaces.
    """
    s = s.lower().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9\s]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

@st.cache_data
def load_rate_database(excel_path: str):
    """
    Iterate through all sheets in pay details.xlsx, find header row (Name / Pay Rate),
    read valid rows, and build:
      - custom_rates: {raw_name -> rate}
      - normalized_rates: {normalize_name(raw_name) -> rate}
      - norm_to_raw: {normalize_name(raw_name) -> raw_name}
    """
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

    normalized_keys = list(normalized_rates.keys())
    return custom_rates, normalized_rates, normalized_keys, norm_to_raw

if st.sidebar.button("🔄 Reload Pay Rates"):
    st.cache_data.clear()
    st.experimental_rerun()

custom_rates, normalized_rates, normalized_keys, norm_to_raw = load_rate_database(RATE_FILE_PATH)

# ──────────────────────────────────────────────────────────────────────────────
# 3) FUZZY LOOKUP & PAY CALCULATION FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────
def lookup_match(name: str):
    """
    Return (matched_raw_name, rate, ratio) for a given extracted name.
    Uses exact normalized match or SequenceMatcher fuzzy match above SIMILARITY_THRESHOLD.
    """
    norm = normalize_name(name)
    if not norm:
        return None, DEFAULT_RATE, 0.0

    if norm in normalized_rates:
        raw = norm_to_raw[norm]
        return raw, normalized_rates[norm], 1.0

    best_ratio = 0.0
    best_norm = None
    for nk in normalized_keys:
        ratio = SequenceMatcher(None, nk, norm).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_norm = nk

    if best_norm and best_ratio >= SIMILARITY_THRESHOLD:
        raw = norm_to_raw[best_norm]
        return raw, normalized_rates[best_norm], best_ratio

    return None, DEFAULT_RATE, best_ratio

def hhmm_to_float(hhmm: str) -> float:
    """
    Convert "HH:MM" string to a float number of hours. Returns 0.0 on failure.
    """
    try:
        h, m = hhmm.strip().split(":")
        return int(h) + int(m) / 60.0
    except:
        return 0.0

def calculate_pay(name: str, daily_data: list[dict]):
    """
    Given an extracted name and a list of daily_data entries (each {'weekday': str, 'hours': float}),
    returns a tuple: (weekday_hours, sat_hours, sun_hours, rate, total_pay, matched_raw, ratio).
    - Days Mon-Fri are summed into weekday_hours, Sat into sat_hours, Sun into sun_hours.
    - Regular weekday pay up to 50 hrs at rate; any beyond at 1.5×.
    - Saturday pay at 1.5×; Sunday pay at 1.75×.
    """
    matched_raw, rate, ratio = lookup_match(name)

    weekday_hours = 0.0
    sat_hours = 0.0
    sun_hours = 0.0

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
    regular_wd = weekday_hours - overtime

    pay_regular = regular_wd * rate
    pay_overtime = overtime * rate * 1.5
    pay_sat = sat_hours * rate * 1.5
    pay_sun = sun_hours * rate * 1.75

    total_pay = pay_regular + pay_overtime + pay_sat + pay_sun
    return weekday_hours, sat_hours, sun_hours, rate, total_pay, matched_raw, ratio

# ──────────────────────────────────────────────────────────────────────────────
# 4) DOCX EXTRACTION FUNCTION
# ──────────────────────────────────────────────────────────────────────────────
def extract_timesheet_data(file) -> dict:
    """
    Parse a .docx timesheet:
      1. Scan all table cells for “Client”; once found, extract the next ALL-CAPS line as Name.
      2. In same table(s), collect daily hours from column index 4 (if numeric), keyed by weekday in column index 1.
      3. Scan for “Site Address: …”.
      4. Fallback: reversed paragraphs for ALL-CAPS name if no name found.
      5. Fallback: filename (stem) as name.
      6. Call calculate_pay(...) to get (weekday_hours, sat_hours, sun_hours, rate, total_pay, matched_raw, ratio).
      7. Return a dict with all fields including “Source File”.
    """
    doc = docx.Document(file)
    name = client = site_address = ""
    date_list = []
    daily_data = []

    # 4.1) Find “Client” cell → client & Name
    for table in doc.tables:
        found_client = False
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text or ""
                if "Client" in txt:
                    found_client = True
                    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
                    for idx, ln in enumerate(lines):
                        if ln.lower().startswith("client"):
                            parts = re.split(r"Client[:\-\s]+", ln, flags=re.IGNORECASE)
                            if len(parts) > 1:
                                client = parts[1].strip()
                            if idx + 1 < len(lines):
                                cand = lines[idx + 1]
                                if cand == cand.upper() and len(cand.split()) >= 2:
                                    name = cand.title()
                            break
                    break
            if found_client:
                break
        if found_client:
            break

    # 4.2) Collect hours, site address, date list
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 5:
                hrs_txt = cells[4].text.strip()
                day_txt = cells[1].text.strip()
                date_txt = cells[0].text.strip()
                if hrs_txt and hrs_txt not in ["-", "–", "—"] and day_txt:
                    try:
                        val = float(hrs_txt)
                        daily_data.append({"weekday": day_txt, "hours": val})
                    except:
                        pass
                if re.match(r"\d{2}\.\d{2}\.\d{4}", date_txt):
                    try:
                        d_obj = datetime.strptime(date_txt, "%d.%m.%Y")
                        date_list.append(d_obj)
                    except:
                        pass
            for cell in cells:
                txt = (cell.text or "").strip()
                if "Site Address" in txt and not site_address:
                    m = re.search(r"Site Address[:\-\s]*(.+)", txt)
                    if m:
                        site_address = m.group(1).strip()

    # 4.3) Fallback: reversed paragraphs for ALL-CAPS name
    if not name:
        for para in reversed(doc.paragraphs):
            text = (para.text or "").strip()
            if text and text == text.upper() and len(text.split()) >= 2 and "PRL" not in text:
                name = text.title()
                break

    # 4.4) Fallback: filename as name
    if not name:
        stem = Path(file.name).stem
        name = stem.replace("_", " ").replace("-", " ").title()

    # 4.5) Calculate pay breakdown
    wd_hrs, sat_hrs, sun_hrs, rate, total_pay, matched_raw, ratio = calculate_pay(name, daily_data)

    # 4.6) Build date_range string if dates exist
    date_range = ""
    if date_list:
        ds = sorted(date_list)
        date_range = f"{ds[0].strftime('%d.%m.%Y')}–{ds[-1].strftime('%d.%m.%Y')}"

    return {
        "Name": name,
        "Matched As": matched_raw or "No match",
        "Ratio": round(ratio, 2),
        "Client": client,
        "Site Address": site_address,
        "Department": "",
        "Weekday Hours": wd_hrs,
        "Saturday Hours": sat_hrs,
        "Sunday Hours": sun_hrs,
        "Rate (£)": rate,
        "Date Range": date_range,
        "Extracted On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Source File": file.name
    }

# ──────────────────────────────────────────────────────────────────────────────
# 5) PDF EXTRACTION FUNCTION
# ──────────────────────────────────────────────────────────────────────────────
def extract_timesheet_data_pdf(file) -> list[dict]:
    """
    Parse an Anel PDF:
      1. Extract “Report Range: … to …” on page 1 for date_range.
      2. Find header line containing “ID Name Paylink” & “Tot Sat-Sun”.
      3. For each subsequent line until “Grand Totals”:
         - tokens = line.split()
         - If last 9 tokens match H:MM, treat as daily_data block.
         - dash_idx = tokens.index("-")
         - Name = tokens[1:dash_idx] → .title()
         - Build daily_data from Mon–Sun slots via hhmm_to_float.
         - Call calculate_pay(name, daily_data).
         - Build a dict with fields including “Weekday Hours”, “Saturday Hours”, “Sunday Hours”.
      4. Return list of dicts (one per person).
    """
    with pdfplumber.open(file) as pdf:
        page1 = pdf.pages[0]
        raw = page1.extract_text()
    lines = raw.split("\n")

    date_range = ""
    for line in lines:
        if line.startswith("Report Range:"):
            m = re.search(
                r"Report Range:\s*(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{2}/\d{2}/\d{2})\s+to\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{2}/\d{2}/\d{2})",
                line
            )
            if m:
                def fmt(d: str) -> str:
                    dt = datetime.strptime(d, "%d/%m/%y")
                    return dt.strftime("%d.%m.%Y")
                date_range = f"{fmt(m.group(1))}–{fmt(m.group(2))}"
            break

    header_idx = None
    for idx, line in enumerate(lines):
        if "ID Name Paylink" in line and "Tot Sat-Sun" in line:
            header_idx = idx
            break

    if header_idx is None:
        return [{
            "Name": "",
            "Matched As": "No match",
            "Ratio": 0.0,
            "Client": "",
            "Site Address": "",
            "Department": "",
            "Weekday Hours": 0.0,
            "Saturday Hours": 0.0,
            "Sunday Hours": 0.0,
            "Rate (£)": 0.0,
            "Date Range": date_range,
            "Extracted On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Source File": file.name
        }]

    results: list[dict] = []
    time_re = re.compile(r"^\d{1,2}:\d{2}$")

    for raw_line in lines[header_idx+1:]:
        if raw_line.strip().startswith("Grand Totals"):
            break
        line = raw_line.strip()
        if not line:
            continue

        tokens = line.split()
        n = len(tokens)
        if n < 10:
            continue

        block = tokens[-9:]
        if not all(time_re.match(t) for t in block):
            continue

        try:
            dash_idx = tokens.index("-")
        except ValueError:
            continue

        name = " ".join(tokens[1:dash_idx]).title()
        daily_data = []
        for wd, idx_tok in [
            ("Monday", 0), ("Tuesday", 1), ("Wednesday", 2),
            ("Thursday", 3), ("Friday", 4), ("Saturday", 6), ("Sunday", 7)
        ]:
            h = hhmm_to_float(block[idx_tok])
            if h > 0:
                daily_data.append({"weekday": wd, "hours": h})

        wd_hrs, sat_hrs, sun_hrs, rate, total_pay, matched_raw, ratio = calculate_pay(name, daily_data)

        comp = ""
        if dash_idx + 2 < n:
            comp = " ".join(tokens[dash_idx+1 : dash_idx+3])

        site_addr = ""
        if dash_idx + 5 < n:
            site_addr = " ".join(tokens[dash_idx+3 : dash_idx+5])

        dept = ""
        dept_start = dash_idx + 5
        dept_end = n - 9
        if dept_start < dept_end:
            dept = " ".join(tokens[dept_start : dept_end]).rstrip("-")

        results.append({
            "Name": name,
            "Matched As": matched_raw or "No match",
            "Ratio": round(ratio, 2),
            "Client": comp,
            "Site Address": site_addr,
            "Department": dept,
            "Weekday Hours": wd_hrs,
            "Saturday Hours": sat_hrs,
            "Sunday Hours": sun_hrs,
            "Rate (£)": rate,
            "Date Range": date_range,
            "Extracted On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Source File": file.name
        })

    if not results:
        return [{
            "Name": "",
            "Matched As": "No match",
            "Ratio": 0.0,
            "Client": "",
            "Site Address": "",
            "Department": "",
            "Weekday Hours": 0.0,
            "Saturday Hours": 0.0,
            "Sunday Hours": 0.0,
            "Rate (£)": 0.0,
            "Date Range": date_range,
            "Extracted On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Source File": file.name
        }]

    return results

# ──────────────────────────────────────────────────────────────────────────────
# 6) MAIN UI WITH TABS FOR “Upload & Review”, “History”, “Dashboard”, “Settings”
# ──────────────────────────────────────────────────────────────────────────────
tabs = st.tabs(["Upload & Review", "History", "Dashboard", "Settings"])

# ──────────────────────────────────────────────────────────────────────────────
# Tab 1: Upload & Review
# ──────────────────────────────────────────────────────────────────────────────
with tabs[0]:
    st.sidebar.header("Upload Timesheets")
    st.sidebar.markdown(
        """
        1. Upload **.docx** or **.pdf** timesheet files.  
        2. Confirm name‐match (expand Debug).  
        3. Override any “No match” names.  
        4. Export final Excel with formulas.
        """
    )

    uploaded_files = st.sidebar.file_uploader(
        "Select files",
        type=["docx", "pdf"],
        accept_multiple_files=True
    )

    if uploaded_files:
        all_rows = []
        progress = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            lower = file.name.lower()
            if lower.endswith(".docx"):
                rec = extract_timesheet_data(file)
                if not rec["Name"]:
                    stem = Path(file.name).stem
                    rec["Name"] = stem.replace("_", " ").replace("-", " ").title()
                all_rows.append(rec)
            elif lower.endswith(".pdf"):
                pdf_list = extract_timesheet_data_pdf(file)
                for r in pdf_list:
                    if not r["Name"]:
                        continue
                    all_rows.append(r)
            progress.progress((i + 1) / total_files)

        df = pd.DataFrame(all_rows)

        st.markdown("### 🔎 Debug: Extracted vs. Matched Pay-Detail Entries")
        debug_df = (
            df[["Name", "Matched As", "Ratio", "Rate (£)", "Source File"]]
            .drop_duplicates()
            .reset_index(drop=True)
        )

        def highlight_low_ratio(val):
            if isinstance(val, float) and val < SIMILARITY_THRESHOLD:
                return "background-color: #FFCCCC"
            return ""

        styled = debug_df.style.applymap(highlight_low_ratio, subset=["Ratio"])
        with st.expander("Show Name-Match Debug Table"):
            st.dataframe(styled, use_container_width=True)

        problem_rows = []
        for idx, row in df.iterrows():
            if row["Weekday Hours"] < 0 or row["Saturday Hours"] < 0 or row["Sunday Hours"] < 0:
                problem_rows.append((idx, "Negative hours"))
            if row["Weekday Hours"] > 168:
                problem_rows.append((idx, "Weekday > 168 hrs"))
            if row["Saturday Hours"] > 24 or row["Sunday Hours"] > 24:
                problem_rows.append((idx, "Weekend hours > 24"))

        if problem_rows:
            st.error("⚠️ Data validation issues found:")
            for idx, reason in problem_rows:
                st.write(f"- Row {idx+1}: {reason} (Name: {df.at[idx,'Name']})")

        if not problem_rows:
            for _, row in df.iterrows():
                c.execute("""
                    INSERT INTO timesheet_entries
                    (name, matched_as, ratio, client, site_address, department,
                     weekday_hours, saturday_hours, sunday_hours, rate,
                     date_range, extracted_on, source_file, upload_timestamp)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (
                    row["Name"], row["Matched As"], row["Ratio"], row["Client"],
                    row["Site Address"], row["Department"], row["Weekday Hours"],
                    row["Saturday Hours"], row["Sunday Hours"], row["Rate (£)"],
                    row["Date Range"], row["Extracted On"], row["Source File"],
                    datetime.now().isoformat()
                ))
            conn.commit()
            st.success(f"✅ Inserted {len(df)} rows into history (SQLite).")

        st.markdown("### 📋 Final Timesheet Table (Read‐Only)")
        if "Calculated Pay (£)" not in df.columns:
            df["Calculated Pay (£)"] = df.apply(
                lambda row: calculate_pay(
                    row["Name"], []
                )[-3],  # total_pay
                axis=1
            )
        st.dataframe(df, use_container_width=True)

        summary_df = (
            df.groupby("Matched As")[["Calculated Pay (£)", "Weekday Hours", "Saturday Hours", "Sunday Hours"]]
            .sum()
            .reset_index()
            .rename(columns={"Matched As": "Name"})
        )

        st.markdown("### 💰 Weekly Summary")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.dataframe(summary_df, use_container_width=True)
        with col2:
            total_hours_all = (
                summary_df["Weekday Hours"].sum()
                + summary_df["Saturday Hours"].sum()
                + summary_df["Sunday Hours"].sum()
            )
            total_pay_all = summary_df["Calculated Pay (£)"].sum()
            st.metric(label="Total Hours This Period", value=f"{total_hours_all:.2f}")
            st.metric(label="Total Pay This Period", value=f"£{total_pay_all:.2f}")

        st.markdown("---")
        st.markdown("### 📥 Download Final Report (Excel with Formulas)")
        output = BytesIO()
        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "Timesheets"

        headers = [
            "Name", "Matched As", "Ratio", "Client", "Site Address", "Department",
            "Weekday Hours", "Saturday Hours", "Sunday Hours", "Rate (£)",
            "Regular Pay (£)", "Overtime Pay (£)", "Saturday Pay (£)", "Sunday Pay (£)", "Total Pay (£)",
            "Date Range", "Extracted On", "Source File"
        ]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        for idx, row in df.iterrows():
            excel_row = idx + 2
            name = row["Name"]
            matched = row["Matched As"]
            ratio = row["Ratio"]
            client = row["Client"]
            site_address = row["Site Address"]
            dept = row["Department"]
            wd_hours = row["Weekday Hours"]
            sat_hours = row["Saturday Hours"]
            sun_hours = row["Sunday Hours"]
            rate = row["Rate (£)"]
            date_range = row["Date Range"]
            extracted_on = row["Extracted On"]
            source_file = row["Source File"]

            col_wd = f"G{excel_row}"
            col_sat = f"H{excel_row}"
            col_sun = f"I{excel_row}"
            col_rate = f"J{excel_row}"

            reg_formula = f"=MIN({col_wd},50)*{col_rate}"
            ot_formula = f"=MAX({col_wd}-50,0)*{col_rate}*1.5"
            sat_formula = f"={col_sat}*{col_rate}*1.5"
            sun_formula = f"={col_sun}*{col_rate}*1.75"
            tot_formula = f"=K{excel_row}+L{excel_row}+M{excel_row}+N{excel_row}"

            ws.append([
                name, matched, ratio, client, site_address, dept,
                wd_hours, sat_hours, sun_hours, rate,
                reg_formula, ot_formula, sat_formula, sun_formula, tot_formula,
                date_range, extracted_on, source_file
            ])

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    val_str = str(cell.value)
                    if len(val_str) > max_length:
                        max_length = len(val_str)
            adjusted_width = (max_length + 2)
            ws.column_dimensions[col_letter].width = adjusted_width

        wb_out.save(output)
        st.download_button(
            "Download Excel with Formulas",
            data=output.getvalue(),
            file_name="PRL_Timesheet_Report_With_Formulas.xlsx",
            help="Excel contains formulas for payroll calculations."
        )

# ──────────────────────────────────────────────────────────────────────────────
# Tab 2: History
# ──────────────────────────────────────────────────────────────────────────────
with tabs[1]:
    st.header("📜 Timesheet History & Audit Trail")

    min_date = st.date_input("From", value=(datetime.today() - timedelta(days=30)).date())
    max_date = st.date_input("To", value=datetime.today().date())
    name_filter = st.text_input("Contractor Name Contains…", "")
    client_filter = st.text_input("Client Contains…", "")

    query = "SELECT * FROM timesheet_entries WHERE upload_timestamp BETWEEN ? AND ?"
    params = [f"{min_date}T00:00:00", f"{max_date}T23:59:59"]
    if name_filter:
        query += " AND name LIKE ?"
        params.append(f"%{name_filter}%")
    if client_filter:
        query += " AND client LIKE ?"
        params.append(f"%{client_filter}%")

    history_df = pd.read_sql_query(query, conn, params=params)
    st.dataframe(history_df, use_container_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# Tab 3: Dashboard
# ──────────────────────────────────────────────────────────────────────────────
with tabs[2]:
    st.header("📊 Dashboard & Analytics")

    df_history = pd.read_sql_query(
        """
        SELECT DATE(upload_timestamp) as date,
               SUM(weekday_hours*rate + saturday_hours*rate*1.5 + sunday_hours*rate*1.75) as total_pay
        FROM timesheet_entries
        GROUP BY DATE(upload_timestamp)
        """, conn)
    if not df_history.empty:
        df_history["date"] = pd.to_datetime(df_history["date"])
        plt.figure(figsize=(8, 4))
        plt.plot(df_history["date"], df_history["total_pay"], marker="o")
        plt.title("Total Pay by Day")
        plt.xlabel("Date")
        plt.ylabel("Total Pay (£)")
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(plt.gcf())
        plt.clf()
    else:
        st.info("No historical data available for trend.")

    st.markdown("---")

    this_month = datetime.today().strftime("%Y-%m")
    qry_top5 = f"""
      SELECT matched_as AS name,
             SUM(weekday_hours*rate + saturday_hours*rate*1.5 + sunday_hours*rate*1.75) AS month_pay
      FROM timesheet_entries
      WHERE strftime('%Y-%m', upload_timestamp) = '{this_month}'
      GROUP BY matched_as
      ORDER BY month_pay DESC
      LIMIT 5
    """
    df_top5 = pd.read_sql_query(qry_top5, conn)
    if not df_top5.empty:
        plt.figure(figsize=(8, 4))
        plt.bar(df_top5["name"], df_top5["month_pay"], color="#4C78A8")
        plt.title("Top 5 Contractors by Pay (This Month)")
        plt.xlabel("Contractor")
        plt.ylabel("Pay (£)")
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(plt.gcf())
        plt.clf()
    else:
        st.info("No entries for this month yet.")

    st.markdown("---")

    df_client = pd.read_sql_query(
        """
        SELECT client, SUM(weekday_hours + saturday_hours + sunday_hours) AS total_hrs
        FROM timesheet_entries
        WHERE upload_timestamp >= date('now','-30 day')
        GROUP BY client
        """, conn)
    if not df_client.empty:
        plt.figure(figsize=(6, 6))
        plt.pie(df_client["total_hrs"], labels=df_client["client"], autopct="%1.1f%%", startangle=140)
        plt.title("Hours by Client (Last 30 Days)")
        plt.tight_layout()
        st.pyplot(plt.gcf())
        plt.clf()
    else:
        st.info("No entries in the last 30 days.")

    st.markdown("---")

    df_ot = pd.read_sql_query(
        """
        SELECT COUNT(*) AS ot_weeks
        FROM (
          SELECT matched_as, strftime('%Y-%W', upload_timestamp) AS week,
                 SUM(weekday_hours) AS total_wd
          FROM timesheet_entries
          GROUP BY matched_as, week
          HAVING total_wd > 50
        )
        """, conn)
    ot_weeks_count = int(df_ot.at[0, "ot_weeks"]) if not df_ot.empty else 0
    st.metric("Number of Contractor-Weeks > 50 Weekday Hours", value=f"{ot_weeks_count}")

# ──────────────────────────────────────────────────────────────────────────────
# Tab 4: Settings
# ──────────────────────────────────────────────────────────────────────────────
with tabs[3]:
    st.header("⚙️ Settings & Configuration")

    st.markdown("### Adjust Fuzzy‐Match Threshold")
    new_thresh = st.slider("Similarity threshold for name matching",
                           min_value=0.50, max_value=0.90,
                           value=SIMILARITY_THRESHOLD, step=0.05)
    if new_thresh != SIMILARITY_THRESHOLD:
        SIMILARITY_THRESHOLD = new_thresh
        st.success(f"Threshold updated to {SIMILARITY_THRESHOLD:.2f} (restart may be required)")

    st.markdown("### Reload Pay Rates")
    if st.button("Reload pay details.xlsx"):
        st.cache_data.clear()
        st.experimental_rerun()

    st.markdown("### Database Info")
    db_size = Path("prl_timesheets.db").stat().st_size / (1024 * 1024)
    st.write(f"📦 SQLite DB file size: {db_size:.2f} MB")

    st.info("Authentication has been disabled for this deployment. Ensure file access is restricted appropriately.")
