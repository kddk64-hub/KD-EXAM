import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from io import BytesIO

st.set_page_config(page_title="DPR → D Form", layout="wide")
st.title("Excel Processing App")
st.caption("Trial Version 2 — DPR Formatting → D Form")

def workbook_to_bytes(wb):
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

def find_header_row(ws, names, max_rows=20):
    wanted = {str(x).strip().lower() for x in names}
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals = {
            str(ws.cell(r, c).value).strip().lower()
            for c in range(1, ws.max_column + 1)
            if ws.cell(r, c).value is not None
        }
        if wanted.issubset(vals):
            return r
    return None

def norm_time(v):
    if v is None:
        return ""
    return str(v).strip().lower().replace("–", "-").replace("—", "-")

def session_count(times):
    u = {norm_time(t) for t in times if norm_time(t)}
    pairs = [
        {"9-10:30 am", "9-11 am"},
        {"12-1:30 pm", "12-2 pm"},
        {"3-4:30 pm", "3-5 pm"},
    ]
    used = set()
    count = 0
    for pair in pairs:
        if pair.issubset(u):
            count += 1
            used.update(pair)
    return count + len(u - used)

def process_dpr(wb):
    # This prototype targets the supplied workbook structure.
    if " bundle all" in wb.sheetnames:
        ws = wb[" bundle all"]
        date_col, time_col = 1, 2
    else:
        ws = None
        for candidate in wb.worksheets:
            h = find_header_row(candidate, {"date", "time"})
            if h:
                ws = candidate
                date_col = next(c for c in range(1,candidate.max_column+1)
                                 if str(candidate.cell(h,c).value).strip().lower()=="date")
                time_col = next(c for c in range(1,candidate.max_column+1)
                                 if str(candidate.cell(h,c).value).strip().lower()=="time")
                break
        if ws is None:
            raise ValueError("Could not find Date/Time columns.")

    # Actual supplied workbook: Date/Time values are merged/carry-forward.
    grouped = {}
    current_date = None
    for r in range(2, ws.max_row + 1):
        d = ws.cell(r, date_col).value
        if d is not None:
            current_date = d
            grouped.setdefault(current_date, [])
        t = ws.cell(r, time_col).value
        if current_date is not None and t is not None:
            grouped[current_date].append(t)

    sessions = {d: session_count(ts) for d, ts in grouped.items()}

    # IMPORTANT: Sessions column is intentionally kept BLANK.
    # It is not populated with dates or calculated values.

    if "Session Summary" in wb.sheetnames:
        del wb["Session Summary"]
    summary = wb.create_sheet("Session Summary")
    summary.append(["Session Name", "Dates", "Total Dates", "Actual Sessions", "Rs."])

    cats = {i: [] for i in range(1, 8)}
    for d, n in sessions.items():
        if 1 <= n <= 7:
            cats[n].append(d)

    for n in range(1, 8):
        ds = sorted(cats[n])
        dates = ", ".join(
            d.strftime("%d/%m") if hasattr(d, "strftime") else str(d)
            for d in ds
        )
        total_dates = len(ds)
        actual = n * total_dates
        summary.append([n, dates, total_dates, actual, actual * 105])

    total_dates = len(sessions)
    total_sessions = sum(sessions.values())
    summary.append(["TOTAL", "", total_dates, total_sessions, total_sessions * 105])
    return wb

def process_d_form(wb):
    if " bundle all" in wb.sheetnames:
        ws = wb[" bundle all"]
        student_col, total_col, block_col, super_col, squad_col = 9,10,11,12,13
        header_row = 1
    else:
        ws = None
        header_row = None
        for candidate in wb.worksheets:
            h = find_header_row(candidate, {"student", "total", "block", "super", "squad"})
            if h:
                ws = candidate
                header_row = h
                break
        if ws is None:
            raise ValueError("Could not find D Form columns.")

        headers = {
            str(ws.cell(header_row,c).value).strip().lower(): c
            for c in range(1, ws.max_column+1)
            if ws.cell(header_row,c).value is not None
        }
        student_col = headers["student"]
        total_col = headers["total"]
        block_col = headers["block"]
        super_col = headers["super"]
        squad_col = headers["squad"]

    # Total: preserve merged structure.
    merged_total_rows = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_col == total_col and rng.max_col == total_col:
            merged_total_rows.update(range(rng.min_row, rng.max_row+1))
            values = []
            for r in range(rng.min_row, rng.max_row+1):
                v = ws.cell(r, student_col).value
                if isinstance(v, (int,float)):
                    values.append(v)
            ws.cell(rng.min_row, total_col).value = sum(values)

    for r in range(header_row+1, ws.max_row+1):
        if r not in merged_total_rows and ws.cell(r,total_col).__class__.__name__ != "MergedCell":
            ws.cell(r,total_col).value = ws.cell(r,student_col).value

    for r in range(header_row+1, ws.max_row+1):
        total = ws.cell(r,total_col).value
        if not isinstance(total,(int,float)):
            continue

        if 0 <= total <= 35: block = 1
        elif 36 <= total <= 70: block = 2
        elif 71 <= total <= 105: block = 3
        elif 106 <= total <= 135: block = 4
        elif 136 <= total <= 165: block = 5
        elif 166 <= total <= 195: block = 6
        elif 196 <= total <= 225: block = 7
        else: continue

        ws.cell(r,block_col).value = block
        ws.cell(r,super_col).value = block
        ws.cell(r,squad_col).value = 0 if block <= 2 else (2 if block <= 4 else 3)

        for c in (total_col, block_col, super_col, squad_col):
            ws.cell(r,c).alignment = Alignment(horizontal="center", vertical="center")

    return wb

if "dpr_output" not in st.session_state:
    st.session_state.dpr_output = None
if "dform_output" not in st.session_state:
    st.session_state.dform_output = None

tab_dpr, tab_dform = st.tabs(["DPR Formatting", "D Form"])

with tab_dpr:
    st.subheader("DPR Formatting")
    uploaded = st.file_uploader("Upload original Excel file", type=["xlsx","xlsm"], key="upload")
    if uploaded and st.button("Process DPR Formatting", type="primary"):
        try:
            wb = load_workbook(uploaded, data_only=False)
            wb = process_dpr(wb)
            st.session_state.dpr_output = workbook_to_bytes(wb)
            st.session_state.dform_output = None
            st.success("DPR Formatting completed. Sessions column remains blank.")
        except Exception as e:
            st.error(str(e))

    if st.session_state.dpr_output:
        st.download_button(
            "Download DPR Processed Excel",
            st.session_state.dpr_output,
            "DPR_Processed.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab_dform:
    st.subheader("D Form")
    if st.session_state.dpr_output:
        st.info("D Form automatically uses the latest DPR Formatting output.")
        if st.button("Process D Form", type="primary"):
            try:
                wb = load_workbook(BytesIO(st.session_state.dpr_output), data_only=False)
                wb = process_d_form(wb)
                st.session_state.dform_output = workbook_to_bytes(wb)
                st.success("D Form processing completed.")
            except Exception as e:
                st.error(str(e))

        if st.session_state.dform_output:
            st.download_button(
                "Download Final Excel",
                st.session_state.dform_output,
                "Final_DPR_DForm.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("First process an Excel file in DPR Formatting.")
