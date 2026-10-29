
import sqlite3
from pathlib import Path
from datetime import datetime, date
import io
import re

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
DB_PATH = DATA_DIR / "gst_reco.db"

for folder in (DATA_DIR, UPLOAD_DIR, REPORT_DIR):
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title="GST Reconciliation System", page_icon="📊", layout="wide")

RETURN_TYPES = {
    "GSTR-2A": "purchase",
    "GSTR-2B": "purchase",
    "GSTR-1": "sales",
    "GSTR-3B": "summary",
}
FIELDS = ["GSTIN", "Party Name", "Invoice No", "Invoice Date", "Invoice Type", "Taxable Value", "IGST", "CGST", "SGST", "Cess", "Total Tax"]
    
ALIASES = {    "Party Name": ["party name", "supplier name", "customer name", "vendor name", "trade name", "legal name", "name of supplier", "name of recipient"],
    "Invoice Type": ["invoice type", "document type", "doc type", "voucher type", "supply type", "type of invoice"],
    "GSTIN": ["gstin", "gstin of supplier", "supplier gstin", "gstin of recipient", "recipient gstin", "ctin", "customer gstin", "party gstin", "gst number"],
    "Invoice No": ["invoice no", "invoice number", "inum", "bill no", "bill number", "document number", "doc no", "voucher no", "voucher number", "inv no", "invoice"],
    "Invoice Date": ["invoice date", "date", "idt", "bill date", "document date", "voucher date"],
    "Taxable Value": ["taxable value", "taxable amount", "taxable", "assessable value", "taxable amt"],
    "IGST": ["igst", "integrated tax", "igst amount", "integrated tax amount"],
    "CGST": ["cgst", "central tax", "cgst amount", "central tax amount"],
    "SGST": ["sgst", "utgst", "state tax", "sgst amount", "state tax amount", "utgst amount"],
    "Cess": ["cess", "cess amount"],
    "Total Tax": ["total tax", "tax amount", "total gst", "gst amount", "total tax amount"],
}

def init_db():
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS import_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            gstin TEXT, period TEXT, return_type TEXT, filename TEXT,
            row_count INTEGER, imported_at TEXT
        )""")
        con.execute("""
        CREATE TABLE IF NOT EXISTS saved_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            gstin TEXT, period TEXT, return_type TEXT, source_name TEXT,
            invoice_key TEXT, invoice_no TEXT, invoice_date TEXT, party_gstin TEXT,
            taxable REAL, igst REAL, cgst REAL, sgst REAL, cess REAL, total_tax REAL,
            imported_at TEXT
        )""")

def norm_col(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()

def read_uploaded_file(uploaded):
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith(".csv"):
        # Try common encodings and separators through pandas' CSV parser.
        for enc in ("utf-8-sig", "cp1252", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except Exception:
                pass
        raise ValueError("CSV file could not be read. Save it as UTF-8 CSV or Excel.")
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(raw), engine="openpyxl")
    if name.endswith(".xls"):
        raise ValueError("Old .xls files are not supported by this app version. In Excel, use Save As → Excel Workbook (*.xlsx) or CSV.")
    raise ValueError("Unsupported file type. Upload .xlsx or .csv.")

def auto_map_columns(df):
    mapped = {}
    normalized = {col: norm_col(col) for col in df.columns}
    for field, aliases in ALIASES.items():
        alias_norms = [norm_col(a) for a in aliases]
        found = None
        for col, nc in normalized.items():
            if nc in alias_norms:
                found = col
                break
        if found is None:
            for col, nc in normalized.items():
                if any(a and (a in nc or nc in a) for a in alias_norms if len(a) >= 4):
                    found = col
                    break
        mapped[field] = found if found is not None else "(Not mapped)"
    return mapped

def numeric(series):
    if series is None:
        return pd.Series(dtype="float64")
    s = series.astype(str).str.replace(",", "", regex=False).str.replace("₹", "", regex=False).str.strip()
    s = s.replace({"": None, "nan": None, "None": None, "-": None})
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def normalize_invoice(value):
    if pd.isna(value):
        return ""
    # Normalize spaces/case but preserve punctuation because invoice references can differ materially.
    return re.sub(r"\s+", "", str(value)).strip().upper()

def standardize(df, mapping, gstin, period, return_type, source_name):
    out = pd.DataFrame()
    for field in FIELDS:
        col = mapping.get(field, "(Not mapped)")
        if col == "(Not mapped)" or col not in df.columns:
            out[field] = "" if field in ("GSTIN", "Party Name", "Invoice No", "Invoice Date", "Invoice Type") else 0.0
        else:
            out[field] = df[col]
    for field in ["Taxable Value", "IGST", "CGST", "SGST", "Cess", "Total Tax"]:
        out[field] = numeric(out[field])
    out["GSTIN"] = out["GSTIN"].fillna("").astype(str).str.strip().str.upper()
    out["Invoice No"] = out["Invoice No"].fillna("").astype(str).str.strip()
    out["Invoice Date"] = pd.to_datetime(out["Invoice Date"], errors="coerce", dayfirst=True).dt.strftime("%Y-%m-%d").fillna("")
    tax_sum = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"]
    out.loc[(out["Total Tax"] == 0) & (tax_sum != 0), "Total Tax"] = tax_sum
    out["GSTIN"] = out["GSTIN"].replace({"nan": "", "None": ""})
    out["Invoice No"] = out["Invoice No"].replace({"nan": "", "None": ""})
    out["Invoice Key"] = out.apply(lambda r: normalize_invoice(r["Invoice No"]), axis=1)
    out["App GSTIN"] = gstin
    out["Period"] = period
    out["Return Type"] = return_type
    out["Source File"] = source_name
    return out


def save_import(df, gstin, period, return_type, filename):
    now = datetime.now().isoformat(timespec="seconds")

    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """INSERT INTO import_log
            (gstin, period, return_type, filename, row_count, imported_at)
            VALUES (?,?,?,?,?,?)""",
            (gstin, period, return_type, filename, int(len(df)), now)
        )

        rows = []
        for _, r in df.iterrows():
            rows.append((
                gstin,
                period,
                return_type,
                filename,
                r["Invoice Key"],
                r["Invoice No"],
                r["Invoice Date"],
                r["GSTIN"],
                float(r["Taxable Value"]),
                float(r["IGST"]),
                float(r["CGST"]),
                float(r["SGST"]),
                float(r["Cess"]),
                float(r["Total Tax"]),
                now,
                str(r.get("Party Name", "") or ""),
                str(r.get("Invoice Type", "") or "")
            ))

        con.executemany(
            """INSERT INTO saved_rows
            (gstin, period, return_type, source_name, invoice_key,
             invoice_no, invoice_date, party_gstin, taxable, igst,
             cgst, sgst, cess, total_tax, imported_at,
             party_name, invoice_type)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            rows
        )



def load_rows(gstin, period, return_type):
    with sqlite3.connect(DB_PATH) as con:
        return pd.read_sql_query(
            """SELECT invoice_key AS 'Invoice Key',
            invoice_no AS 'Invoice No',
            invoice_date AS 'Invoice Date',
            party_gstin AS 'GSTIN',
            party_name AS 'Party Name',
            invoice_type AS 'Invoice Type',
            taxable AS 'Taxable Value',
            igst AS IGST,
            cgst AS CGST,
            sgst AS SGST,
            cess AS Cess,
            total_tax AS 'Total Tax',
            source_name AS 'Source File'
            FROM saved_rows
            WHERE gstin=? AND period=? AND return_type=?""",
            con,
            params=(gstin, period, return_type)
        )

import calendar

def get_month_periods(start_date, end_date):
    periods = []
    year, month = start_date.year, start_date.month

    while (year, month) <= (end_date.year, end_date.month):
        periods.append(f"{month:02d}{year}")

        if month == 12:
            year += 1
            month = 1
        else:
            month += 1

    return periods

def load_rows_for_period(gstin, start_date, end_date, return_type):
    # Books are filtered by invoice date.
    # GST portal returns are filtered by stored return period.
    periods = get_month_periods(start_date, end_date)

    if not periods:
        return pd.DataFrame()

    
    placeholders = ",".join(["?"] * len(periods))

    query = """
        SELECT invoice_key AS 'Invoice Key',
               invoice_no AS 'Invoice No',
               invoice_date AS 'Invoice Date',
               party_gstin AS 'GSTIN',
               party_name AS 'Party Name',
               invoice_type AS 'Invoice Type',
               taxable AS 'Taxable Value',
               igst AS IGST,
               cgst AS CGST,
               sgst AS SGST,
               cess AS Cess,
               total_tax AS 'Total Tax',
               source_name AS 'Source File'
        FROM saved_rows
        WHERE gstin=?
          AND return_type=?
    """

    if return_type in ("BOOKS-PURCHASE", "BOOKS-SALES"):
        query += """
          AND invoice_date >= ?
          AND invoice_date <= ?
          AND invoice_date <> ''
        """
        params = [
            gstin,
            return_type,
            start_date.isoformat(),
            end_date.isoformat()
        ]
    else:
        query += f" AND period IN ({placeholders})"
        params = [gstin, return_type] + periods

    with sqlite3.connect(DB_PATH) as con:
        return pd.read_sql_query(query, con, params=params)


def get_fy_dates(financial_year):
    start_year = int(financial_year.split("-")[0])
    return date(start_year, 4, 1), date(start_year + 1, 3, 31)


def get_quarter_dates(financial_year, quarter):
    start_year = int(financial_year.split("-")[0])
    quarter_months = {
        "Q1": (4, 6),
        "Q2": (7, 9),
        "Q3": (10, 12),
        "Q4": (1, 3)
    }

    first_month, last_month = quarter_months[quarter]
    year = start_year + 1 if quarter == "Q4" else start_year
    last_day = calendar.monthrange(year, last_month)[1]

    return date(year, first_month, 1), date(year, last_month, last_day)


def select_reporting_period():
    st.subheader("Reporting Period")

    mode = st.selectbox(
        "Report period type",
        ["Month-wise", "Quarter-wise", "Financial Year", "Custom Date Range"],
        key="report_period_mode"
    )

    fy = st.selectbox(
        "Financial Year",
        ["2025-26", "2026-27", "2027-28"],
        index=1,
        key="report_fy"
    )

    if mode == "Month-wise":
        months = [
            ("April", 4), ("May", 5), ("June", 6),
            ("July", 7), ("August", 8), ("September", 9),
            ("October", 10), ("November", 11), ("December", 12),
            ("January", 1), ("February", 2), ("March", 3)
        ]

        selected = st.selectbox(
            "Month", [m[0] for m in months], key="report_month"
        )
        month = dict(months)[selected]
        year = int(fy.split("-")[0]) + (1 if month <= 3 else 0)

        start = date(year, month, 1)
        end = date(year, month, calendar.monthrange(year, month)[1])

    elif mode == "Quarter-wise":
        quarter = st.selectbox(
            "Quarter", ["Q1", "Q2", "Q3", "Q4"],
            key="report_quarter"
        )
        start, end = get_quarter_dates(fy, quarter)

    elif mode == "Financial Year":
        start, end = get_fy_dates(fy)

    else:
        default_start, default_end = get_fy_dates(fy)
        start = st.date_input(
            "Start date", value=default_start,
            key="report_custom_start"
        )
        end = st.date_input(
            "End date", value=max(start, default_end),
            min_value=start,
            key="report_custom_end"
        )

    st.caption(f"Selected dates: {start:%d-%m-%Y} to {end:%d-%m-%Y}")
    return start, end


def compare_frames(left, right, left_label, right_label, tolerance):
    # Aggregate duplicates on invoice key + party GSTIN, since source files may contain multiple tax rows.
    keys = ["Invoice Key", "GSTIN"]
    value_cols = ["Taxable Value", "IGST", "CGST", "SGST", "Cess", "Total Tax"]
    for df in (left, right):
        for c in keys:
            if c not in df.columns:
                df[c] = ""
        df["Invoice Key"] = df["Invoice Key"].fillna("").astype(str).str.strip().str.upper()
        df["GSTIN"] = df["GSTIN"].fillna("").astype(str).str.strip().str.upper()
    l = left[left["Invoice Key"] != ""].groupby(keys, dropna=False)[value_cols].sum().reset_index()
    r = right[right["Invoice Key"] != ""].groupby(keys, dropna=False)[value_cols].sum().reset_index()
    merged = l.merge(r, on=keys, how="outer", suffixes=(f"_{left_label}", f"_{right_label}"), indicator=True)
    status = []
    for _, row in merged.iterrows():
        if row["_merge"] == "left_only":
            status.append(f"Only in {left_label}")
        elif row["_merge"] == "right_only":
            status.append(f"Only in {right_label}")
        else:
            checks = []
            for c in value_cols:
                a = float(row.get(f"{c}_{left_label}", 0) or 0)
                b = float(row.get(f"{c}_{right_label}", 0) or 0)
                if abs(a - b) > tolerance:
                    checks.append(c)
            status.append("Matched" if not checks else "Value mismatch: " + ", ".join(checks))
    merged["Status"] = status
    return merged.drop(columns=["_merge"])

init_db()

st.title("GST Reconciliation System")
st.caption("Local-first • SQLite history • Excel/CSV import • No GSP credentials stored in this app")

with st.sidebar:
    st.header("Company / Reporting Period")

    gstin = st.text_input(
        "Your GSTIN (15 characters)", max_chars=15
    ).strip().upper()

    period_mode = st.selectbox(
    "Reporting period",
    ["Month-wise", "Quarter-wise", "Financial Year", "Custom Date Range"],
    key="sidebar_period_mode"
)

fy = st.selectbox(
    "Financial Year",
    ["2025-26", "2026-27", "2027-28"],
    index=1,
    key="sidebar_fy"
)

if period_mode == "Month-wise":
    month_options = [
        ("April", 4), ("May", 5), ("June", 6),
        ("July", 7), ("August", 8), ("September", 9),
        ("October", 10), ("November", 11), ("December", 12),
        ("January", 1), ("February", 2), ("March", 3)
    ]
    selected_month = st.selectbox(
        "Month", [m[0] for m in month_options], key="sidebar_month"
    )
    month = dict(month_options)[selected_month]
    year = int(fy.split("-")[0]) + (1 if month <= 3 else 0)
    report_start = date(year, month, 1)
    report_end = date(year, month, calendar.monthrange(year, month)[1])

elif period_mode == "Quarter-wise":
    quarter = st.selectbox(
        "Quarter", ["Q1", "Q2", "Q3", "Q4"], key="sidebar_quarter"
    )
    report_start, report_end = get_quarter_dates(fy, quarter)

elif period_mode == "Financial Year":
    report_start, report_end = get_fy_dates(fy)

else:
    report_start = st.date_input(
        "Start date", value=date(2026, 4, 1), key="sidebar_custom_start"
    )
    report_end = st.date_input(
        "End date", value=date(2027, 3, 31),
        min_value=report_start, key="sidebar_custom_end"
    )

period = f"{report_start.month:02d}{report_start.year}"
st.caption(
    f"Selected: {report_start:%d-%m-%Y} to {report_end:%d-%m-%Y}"
)
st.caption("Import period: MMYYYY. For quarterly/FY reports, import each tax month separately.")
st.divider()
st.caption("This app runs locally on this computer.")

tab_import, tab_reco, tab_summary, tab_history, tab_help = st.tabs([
    "Import files", "Invoice reconciliation", "GSTR-3B summary", "Import history", "Instructions"
])

with tab_import:
    st.subheader("Import a return file or Tally export")
    st.write("Upload a downloaded portal file or Tally purchase/sales register. Choose the return/source type accurately.")
    return_type = st.selectbox("File type / source", [
        "GSTR-2A", "GSTR-2B", "GSTR-1", "Books - Purchase Register",
        "Books - Sales Register", "GSTR-3B"
    ])
    uploaded = st.file_uploader("Choose .xlsx or .csv file", type=["xlsx", "xlsm", "csv"], key="upload")
    if uploaded is not None:
        try:
            raw_df = read_uploaded_file(uploaded)
            st.write(f"Rows detected: **{len(raw_df):,}** | Columns: **{len(raw_df.columns)}**")
            st.dataframe(raw_df.head(10), use_container_width=True)
            if return_type == "Books - Purchase Register":
                db_type = "BOOKS-PURCHASE"
            elif return_type == "Books - Sales Register":
                db_type = "BOOKS-SALES"
            else:
                db_type = return_type
            st.markdown("#### Map your columns")
            st.caption("The app tries to identify common column headings automatically. Check each selection before importing.")
            detected = auto_map_columns(raw_df)
            options = ["(Not mapped)"] + list(raw_df.columns)
            cols = st.columns(3)
            mapping = {}
            for i, field in enumerate(FIELDS):
                with cols[i % 3]:
                    default = detected.get(field, "(Not mapped)")
                    index = options.index(default) if default in options else 0
                    mapping[field] = st.selectbox(field, options, index=index, key=f"map_{field}")
            if return_type == "GSTR-3B":
                st.warning("GSTR-3B is a summary return, not a normal invoice-level file. This import view can store rows, but invoice matching is not appropriate. Use the GSTR-3B summary tab to compare manually entered totals.")
            if st.button("Import into local database", type="primary"):
                if not gstin or len(gstin) != 15:
                    st.error("Enter your own valid 15-character GSTIN in the sidebar.")
                elif not period or len(period) != 6 or not period.isdigit() or not 1 <= int(period[:2]) <= 12:
                    st.error("Enter return period in MMYYYY format, e.g. 092026.")
                elif return_type in ("GSTR-2A", "GSTR-2B", "GSTR-1", "Books - Purchase Register", "Books - Sales Register") and mapping["Invoice No"] == "(Not mapped)":
                    st.error("Invoice No must be mapped for invoice-level reconciliation.")
                else:
                    standardized = standardize(raw_df, mapping, gstin, period, db_type, uploaded.name)
                    save_import(standardized, gstin, period, db_type, uploaded.name)
                    safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", uploaded.name)
                    standardized.to_csv(UPLOAD_DIR / f"{gstin}_{period}_{db_type}_{safe_name}.csv", index=False)
                    st.success(f"Imported {len(standardized):,} rows into the local database.")
        except Exception as e:
            st.error(f"Could not read/import this file: {e}")

with tab_reco:
    st.subheader("Invoice-level reconciliation")
    st.write("Compare GSTR-2A or GSTR-2B with Tally purchase register, or GSTR-1 with Tally sales register.")
    reco_choice = st.selectbox("Reconciliation", [
        "GSTR-2A vs Books Purchase", "GSTR-2B vs Books Purchase", "GSTR-1 vs Books Sales"
    ])
    tolerance = st.number_input("Amount tolerance (₹)", min_value=0.0, max_value=100.0, value=1.0, step=0.5)
    if st.button("Run reconciliation", type="primary"):
       
        if not gstin or len(gstin) != 15:
            st.error("Enter a valid GSTIN in the sidebar first.")

        else:
            if reco_choice == "GSTR-2A vs Books Purchase":
                left_type, right_type = "GSTR-2A", "BOOKS-PURCHASE"
            elif reco_choice == "GSTR-2B vs Books Purchase":
                left_type, right_type = "GSTR-2B", "BOOKS-PURCHASE"
            else:
                left_type, right_type = "GSTR-1", "BOOKS-SALES"

            left = load_rows_for_period(gstin, report_start, report_end, left_type)
            right = load_rows_for_period(gstin, report_start, report_end, right_type)

            if left.empty or right.empty:
                st.warning(f"Both datasets are needed. Found {len(left)} rows for {left_type} and {len(right)} rows for {right_type}.")
            else:
                result = compare_frames(left, right, "Portal", "Books", tolerance)
                st.session_state["last_reco"] = result
                st.session_state["last_reco_name"] = reco_choice
    if "last_reco" in st.session_state:
        result = st.session_state["last_reco"]
        counts = result["Status"].value_counts()
        c1, c2, c3 = st.columns(3)
        c1.metric("Matched", int(counts.get("Matched", 0)))
        c2.metric("Only in Portal", int(sum(v for k, v in counts.items() if k == "Only in Portal")))
        c3.metric("Exceptions", int(len(result) - int(counts.get("Matched", 0))))
        st.dataframe(result, use_container_width=True, height=420)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            result.to_excel(writer, index=False, sheet_name="Reconciliation")
        st.download_button("Download reconciliation report (.xlsx)", data=out.getvalue(),
                           file_name="GST_Reconciliation_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_summary:
    st.subheader("GSTR-3B summary reconciliation")
    st.write("GSTR-3B is a return-level summary. Compare figures with your books and portal summary; this is not invoice matching.")
    st.caption("Enter values from the relevant GSTR-3B table and books summary. Values are not submitted to GST portal.")
    fields = [
        ("Outward taxable value", "outward"),
        ("IGST liability", "igst"),
        ("CGST liability", "cgst"),
        ("SGST/UTGST liability", "sgst"),
        ("Cess liability", "cess"),
        ("Eligible ITC", "itc"),
        ("RCM liability", "rcm"),
    ]
    rows = []
    for label, key in fields:
        c1, c2, c3 = st.columns([2, 1, 1])
        c1.write(label)
        p = c2.number_input(f"Portal {label} (₹)", min_value=0.0, value=0.0, key=f"p_{key}")
        b = c3.number_input(f"Books {label} (₹)", min_value=0.0, value=0.0, key=f"b_{key}")
        rows.append({"Particulars": label, "Portal": p, "Books": b, "Difference (Books - Portal)": round(b-p, 2)})
    summary = pd.DataFrame(rows)
    st.dataframe(summary, use_container_width=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="3B Summary")
    st.download_button("Download GSTR-3B summary comparison", data=out.getvalue(),
                       file_name="GSTR3B_Summary_Reconciliation.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_history:
    st.subheader("Local import history")
    if gstin and len(gstin) == 15:
        with sqlite3.connect(DB_PATH) as con:
            hist = pd.read_sql_query(
                "SELECT period, return_type, filename, row_count, imported_at FROM import_log WHERE gstin=? ORDER BY id DESC LIMIT 200",
                con, params=(gstin,))
        st.dataframe(hist, use_container_width=True)
    else:
        st.info("Enter your GSTIN in the sidebar to see import history.")

with tab_help:
    st.subheader("How to use this version")
    st.markdown("""
1. Export the Tally purchase register and sales register to `.xlsx` or `.csv`.
2. Download GSTR-2A, GSTR-2B and GSTR-1 files from the GST portal, then upload each in **Import files**.
3. Upload Tally purchase register as **Books - Purchase Register** and sales register as **Books - Sales Register**.
4. Review and correct the column mapping before importing.
5. Open **Invoice reconciliation**, choose a comparison, and run it.
6. Export the result to Excel and review exceptions before taking any ITC, tax, or filing decision.
7. In **GSTR-3B summary**, manually enter the portal and books totals and export the comparison.
8. Use the same GSTIN and period for all related files. Period format is MMYYYY, e.g. September 2026 = 092026.

**Important limitations of this starter version**
- It does not connect to a GSP or download returns automatically.
- Portal export formats vary. Column mapping must be checked; complex multi-sheet/JSON exports may need format-specific handling.
- GSTR-3B is summary-level and is compared by totals, not invoice number.
- It uses invoice number + party GSTIN as a basic matching key. Credit/debit notes, amendments, reverse-charge cases, ISD, imports, duplicate invoice numbers and credit notes require more specialized logic.
- A match in this report does not itself establish ITC eligibility or legal compliance. Verify exceptions with the original documents and current rules.
- The database is local. This app does not yet encrypt the database or create encrypted cloud backups. Keep Windows access protected and arrange a tested encrypted backup before storing sensitive data.
""")
