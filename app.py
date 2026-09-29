import re
from copy import copy
from datetime import date, datetime
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="DPR → D Form → Session Count", layout="wide")
st.title("Excel Processing App")
st.caption("Version 5.1 — 3 independent prompts with automatic file handoff")

TARGET_HEADERS = ["Total", "Block", "Super", "Squad", "Sessions"]


def wb_bytes(wb):
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def norm_header(v):
    return re.sub(r"\s+", " ", str(v or "").strip().lower())


def find_header_row(ws, required, max_rows=20):
    req = {norm_header(x) for x in required}
    for r in range(1, min(ws.max_row, max_rows) + 1):
        vals = {norm_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        if req.issubset(vals):
            return r
    return None


def find_col(ws, names, header_row=1):
    wanted = {norm_header(x) for x in names}
    for c in range(1, ws.max_column + 1):
        if norm_header(ws.cell(header_row, c).value) in wanted:
            return c
    return None


def get_main_sheet(wb):
    if " bundle all" in wb.sheetnames:
        return wb[" bundle all"]
    for ws in wb.worksheets:
        h = find_header_row(ws, {"date", "time"})
        if h:
            return ws
    return wb.active


def date_key(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def display_date(k):
    try:
        return datetime.strptime(k, "%Y-%m-%d").strftime("%d/%m")
    except Exception:
        return str(k)


def norm_time(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%I:%M %p").lstrip("0").lower()
    if hasattr(v, "strftime") and not isinstance(v, str):
        try:
            return v.strftime("%I:%M %p").lstrip("0").lower()
        except Exception:
            pass
    s = str(v).strip().lower().replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+to\s+", "-", s)
    s = s.replace(".30", ":30").replace(".00", ":00")
    s = re.sub(r"\s*([:-])\s*", r"\1", s)
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"(\d)(am|pm)$", r"\1 \2", s)
    return s


def copy_cell_style(src, dst):
    if isinstance(dst, MergedCell) or isinstance(src, MergedCell):
        return
    if src.has_style:
        dst._style = copy(src._style)
    dst.number_format = src.number_format
    dst.alignment = copy(src.alignment)
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.protection = copy(src.protection)


def copy_col_dimension(ws, src_col, dst_col):
    s = ws.column_dimensions[get_column_letter(src_col)]
    d = ws.column_dimensions[get_column_letter(dst_col)]
    d.width = s.width
    d.hidden = s.hidden
    d.bestFit = s.bestFit
    d.outlineLevel = s.outlineLevel
    d.collapsed = s.collapsed


def unmerge_ranges_in_column(ws, col):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == col and rng.max_col == col:
            ws.unmerge_cells(str(rng))


def copy_column_with_vertical_merges(ws, src_col, dst_col, clear_destination=True):
    """Copy a single column's values/styles and vertical merge structure."""
    copy_col_dimension(ws, src_col, dst_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, src_col)
        dst = ws.cell(r, dst_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            if clear_destination and r != 1:
                dst.value = None

    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == src_col and rng.max_col == src_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            top = ws.cell(rng.min_row, dst_col)
            copy_cell_style(ws.cell(rng.min_row, src_col), top)
            if clear_destination:
                top.value = None


def copy_vertical_merge_structure(ws, src_col, dst_col, header_row, clear=True):
    copy_col_dimension(ws, src_col, dst_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, src_col)
        dst = ws.cell(r, dst_col)
        copy_cell_style(src, dst)
        if clear and r != header_row and not isinstance(dst, MergedCell):
            dst.value = None
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == src_col and rng.max_col == src_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            top = ws.cell(rng.min_row, dst_col)
            top.value = None
            copy_cell_style(ws.cell(rng.min_row, src_col), top)


def merge_contiguous_date_time(ws, date_col, time_col, header_row):
    unmerge_ranges_in_column(ws, date_col)
    unmerge_ranges_in_column(ws, time_col)

    current = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v is not None and str(v).strip() != "":
            current = v
        elif current is not None:
            ws.cell(r, date_col).value = current

    rows = list(range(header_row + 1, ws.max_row + 1))
    i = 0
    while i < len(rows):
        r0 = rows[i]
        v0 = date_key(ws.cell(r0, date_col).value)
        if not v0:
            i += 1
            continue
        j = i + 1
        while j < len(rows) and date_key(ws.cell(rows[j], date_col).value) == v0:
            j += 1
        r1 = rows[j - 1]
        if r1 > r0:
            ws.merge_cells(start_row=r0, start_column=date_col, end_row=r1, end_column=date_col)
        i = j

    date_blocks = []
    i = header_row + 1
    while i <= ws.max_row:
        dk = date_key(ws.cell(i, date_col).value)
        if not dk:
            i += 1
            continue
        end = i
        for rng in ws.merged_cells.ranges:
            if rng.min_col == date_col and rng.max_col == date_col and rng.min_row == i:
                end = rng.max_row
                break
        date_blocks.append((i, end))
        i = end + 1

    for d0, d1 in date_blocks:
        i = d0
        while i <= d1:
            tv = norm_time(ws.cell(i, time_col).value)
            if not tv:
                i += 1
                continue
            j = i + 1
            while j <= d1 and norm_time(ws.cell(j, time_col).value) == tv:
                j += 1
            if j - 1 > i:
                ws.merge_cells(start_row=i, start_column=time_col,
                               end_row=j - 1, end_column=time_col)
            i = j


def set_alignment(ws, col, horizontal, vertical="center", header_row=None):
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, col)
        if isinstance(cell, MergedCell):
            continue
        cell.alignment = copy(cell.alignment)
        cell.alignment = Alignment(
            horizontal=horizontal,
            vertical=vertical,
            text_rotation=cell.alignment.text_rotation,
            wrap_text=cell.alignment.wrap_text,
            shrink_to_fit=cell.alignment.shrink_to_fit,
            indent=cell.alignment.indent,
        )


def ensure_target_columns(ws, header_row):
    existing = {norm_header(ws.cell(header_row, c).value): c for c in range(1, ws.max_column + 1)}
    found = [existing.get(norm_header(h)) for h in TARGET_HEADERS]
    if all(found):
        return dict(zip(TARGET_HEADERS, found))

    start = ws.max_column + 1
    date_col = find_col(ws, {"date"}, header_row)
    time_col = find_col(ws, {"time"}, header_row)
    if not date_col or not time_col:
        raise ValueError("Date and Time columns are required for DPR Formatting.")

    cols = {}
    for idx, h in enumerate(TARGET_HEADERS):
        c = start + idx
        src_col = date_col if h == "Sessions" else time_col
        copy_vertical_merge_structure(ws, src_col, c, header_row, clear=True)
        ws.cell(header_row, c).value = h
        cols[h] = c
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            cell.font = f
    return cols


def process_formatting(wb):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row) or 1
    time_col = find_col(ws, {"time"}, header_row) or 2

    merge_contiguous_date_time(ws, date_col, time_col, header_row)
    cols = ensure_target_columns(ws, header_row)

    # Final requested alignment for original Date/Time columns.
    set_alignment(ws, date_col, "left", "center")
    set_alignment(ws, time_col, "right", "center")

    # New calculation columns are centered; Sessions remains completely blank.
    for h, c in cols.items():
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            cell.font = f
            if h == "Sessions":
                cell.value = None

    return wb


def student_col_for(ws, header_row):
    c = find_col(ws, {"student", "stu.no", "stu no", "student no", "student number"}, header_row)
    if c:
        return c
    if ws.max_column >= 9:
        return 9
    raise ValueError("Could not find the Student column.")


def numeric(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        try:
            return float(s)
        except Exception:
            return None
    return None



def copy_date_format_and_merges(ws, date_col, dst_col, header_row, value):
    """Copy Date-column formatting/vertical merge structure to a new D Form column.
    Populate each visible (unmerged or merged top-left) cell with the requested value.
    """
    unmerge_ranges_in_column(ws, dst_col)
    copy_col_dimension(ws, date_col, dst_col)

    # Copy the Date column's cell formatting exactly.
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, date_col)
        dst = ws.cell(r, dst_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)

    # Recreate the Date column's vertical merged structure.
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == date_col and rng.max_col == date_col:
            ws.merge_cells(start_row=rng.min_row, start_column=dst_col,
                           end_row=rng.max_row, end_column=dst_col)
            copy_cell_style(ws.cell(rng.min_row, date_col), ws.cell(rng.min_row, dst_col))

    # Header keeps its own heading; all data/visible cells receive the requested value.
    ws.cell(header_row, dst_col).alignment = copy(ws.cell(header_row, date_col).alignment)
    ws.cell(header_row, dst_col).value = None
    for r in range(header_row + 1, ws.max_row + 1):
        cell = ws.cell(r, dst_col)
        if isinstance(cell, MergedCell):
            continue
        cell.value = value


def move_sessions_to_final_position(ws, header_row, sessions_col):
    """Make final D Form tail contiguous: ... Int. Squad | Int. Sr. Superior | Ext. Sr. Supervisor | Sessions.
    The original Sessions column is copied to the new final position, preserving its vertical merge structure.
    """
    # Reserve two contiguous new columns immediately after the old Sessions column.
    # Final layout is: Int. Sr. Superior | Ext. Sr. Supervisor | Sessions.
    final_sessions_col = max(ws.max_column + 2, sessions_col + 2)
    # Copy original Sessions structure to the new final column before clearing the source.
    unmerge_ranges_in_column(ws, final_sessions_col)
    copy_col_dimension(ws, sessions_col, final_sessions_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, sessions_col)
        dst = ws.cell(r, final_sessions_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            dst.value = src.value
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == sessions_col and rng.max_col == sessions_col:
            ws.merge_cells(start_row=rng.min_row, start_column=final_sessions_col,
                           end_row=rng.max_row, end_column=final_sessions_col)
            top = ws.cell(rng.min_row, final_sessions_col)
            src_top = ws.cell(rng.min_row, sessions_col)
            copy_cell_style(src_top, top)
            top.value = src_top.value

    # Remove the old Sessions merge structure, then use it as Int. Sr. Superior.
    unmerge_ranges_in_column(ws, sessions_col)
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, sessions_col)
        if not isinstance(cell, MergedCell):
            cell.value = None
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font); f.bold = True; cell.font = f

    # The new Ext. Sr. Supervisor column is inserted logically by using the next column.
    ext_col = sessions_col + 1
    if ext_col != final_sessions_col:
        # There should be no unrelated column between the old Sessions and appended final Sessions.
        # Copy any current content/styles only if necessary, then clear it for the new column.
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, ext_col)
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.alignment = Alignment(horizontal="center", vertical="center")
        copy_col_dimension(ws, sessions_col, ext_col)

    ws.cell(header_row, sessions_col).value = "Int. Sr. Superior"
    ws.cell(header_row, ext_col).value = "Ext. Sr. Supervisor"
    ws.cell(header_row, final_sessions_col).value = "Sessions"

    for c in (sessions_col, ext_col, final_sessions_col):
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font); f.bold = True; cell.font = f
            if c == final_sessions_col:
                cell.value = None

    return final_sessions_col


def process_dform(wb):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"total", "block", "super", "squad"}) or 1
    student_col = student_col_for(ws, header_row)
    total_col = find_col(ws, {"total"}, header_row)
    block_col = find_col(ws, {"block"}, header_row)
    super_col = find_col(ws, {"super"}, header_row)
    squad_col = find_col(ws, {"squad"}, header_row)
    sessions_col = find_col(ws, {"sessions"}, header_row)
    if not all([total_col, block_col, super_col, squad_col, sessions_col]):
        raise ValueError("DPR Formatting must create Total, Block, Super, Squad and Sessions columns first.")

    merged_total_rows = set()
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == total_col and rng.max_col == total_col and rng.min_row > header_row:
            merged_total_rows.update(range(rng.min_row, rng.max_row + 1))
            total = 0
            found = False
            for r in range(rng.min_row, rng.max_row + 1):
                n = numeric(ws.cell(r, student_col).value)
                if n is not None:
                    total += n
                    found = True
            top = ws.cell(rng.min_row, total_col)
            if not isinstance(top, MergedCell):
                top.value = total if found else 0

    for r in range(header_row + 1, ws.max_row + 1):
        tc = ws.cell(r, total_col)
        if r not in merged_total_rows and not isinstance(tc, MergedCell):
            tc.value = ws.cell(r, student_col).value

    for r in range(header_row + 1, ws.max_row + 1):
        total = numeric(ws.cell(r, total_col).value)
        if total is None:
            continue
        if 0 <= total <= 35:
            block = 1
        elif 36 <= total <= 70:
            block = 2
        elif 71 <= total <= 105:
            block = 3
        elif 106 <= total <= 135:
            block = 4
        elif 136 <= total <= 165:
            block = 5
        elif 166 <= total <= 195:
            block = 6
        elif 196 <= total <= 225:
            block = 7
        else:
            continue
        squad = 0 if block <= 2 else (2 if block <= 4 else 3)
        for c, value in ((block_col, block), (super_col, block), (squad_col, squad)):
            cell = ws.cell(r, c)
            if not isinstance(cell, MergedCell):
                cell.value = value
                cell.alignment = Alignment(horizontal="center", vertical="center")
                f = copy(cell.font); f.bold = True; cell.font = f

    # Requested D Form headings.
    ws.cell(header_row, total_col).value = "Total Students"
    ws.cell(header_row, block_col).value = "No. of Blocks"
    ws.cell(header_row, super_col).value = "Supervisors"
    ws.cell(header_row, squad_col).value = "Int. Squad"

    # Keep the existing Sessions column blank, but move it to the final contiguous position and
    # use the old position for the first new blank column.
    final_sessions_col = move_sessions_to_final_position(ws, header_row, sessions_col)

    # Second new column is immediately after the first; no blank gap is allowed.
    int_sr_col = sessions_col
    ext_sr_col = sessions_col + 1
    # Int. Sr. Supervisor and Ext. Sr. Supervisor must look exactly like the Date column.
    # Their visible cells are populated with 1 and 2 respectively.
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column is required for the supervisor columns.")
    copy_date_format_and_merges(ws, date_col, int_sr_col, header_row, 1)
    copy_date_format_and_merges(ws, date_col, ext_sr_col, header_row, 2)

    # Requested numeric formatting:
    # Total Students through Ext. Sr. Supervisor must be centered horizontally/vertically,
    # bold, and slightly larger. Sessions remains blank and is not included here.
    numeric_format_cols = (
        total_col, block_col, super_col, squad_col, int_sr_col, ext_sr_col
    )
    for c in numeric_format_cols:
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            base_size = f.sz if f.sz else 11
            f.sz = max(base_size + 1, 12)
            cell.font = f

    return wb


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


def apply_summary_borders(ws):
    side = Side(style="thin")
    border = Border(left=side, right=side, top=side, bottom=side)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def process_sessions(wb):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    time_col = find_col(ws, {"time"}, header_row)
    if not date_col or not time_col:
        raise ValueError("Date and Time columns are required for Session Count.")

    by_date = {}
    current_date = ""
    for r in range(header_row + 1, ws.max_row + 1):
        d = ws.cell(r, date_col).value
        if d is not None and str(d).strip() != "":
            current_date = date_key(d)
        t = ws.cell(r, time_col).value
        if current_date and t is not None and str(t).strip() != "":
            by_date.setdefault(current_date, []).append(t)

    counts = {d: session_count(ts) for d, ts in by_date.items() if d}

    sessions_col = find_col(ws, {"sessions"}, header_row)
    if sessions_col:
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, sessions_col)
            if not isinstance(cell, MergedCell):
                cell.value = None

    if "Session Summary" in wb.sheetnames:
        del wb["Session Summary"]
    summary = wb.create_sheet("Session Summary")
    headers = ["Session Name", "Dates", "Total Dates", "Actual Sessions", "Rs."]
    summary.append(headers)
    for c in range(1, 6):
        summary.cell(1, c).font = Font(bold=True)
        summary.cell(1, c).alignment = Alignment(horizontal="center", vertical="center")

    cats = {i: [] for i in range(1, 8)}
    for d, n in counts.items():
        if 1 <= n <= 7:
            cats[n].append(d)

    for n in range(1, 8):
        ds = sorted(cats[n])
        total_dates = len(ds)
        actual = n * total_dates
        summary.append([n, ", ".join(display_date(d) for d in ds), total_dates, actual, actual * 105])

    total_dates = len(counts)
    total_sessions = sum(counts.values())
    summary.append(["TOTAL", "", total_dates, total_sessions, total_sessions * 105])

    independent = sum(session_count(ts) for ts in by_date.values())
    if independent != total_sessions:
        raise ValueError(f"Session verification failed: summary={total_sessions}, independent={independent}")

    summary.column_dimensions["A"].width = 15
    summary.column_dimensions["B"].width = 55
    summary.column_dimensions["C"].width = 15
    summary.column_dimensions["D"].width = 18
    summary.column_dimensions["E"].width = 15
    apply_summary_borders(summary)

    return wb, total_dates, total_sessions


for key in ("original_bytes", "format_bytes", "dform_bytes", "final_bytes"):
    if key not in st.session_state:
        st.session_state[key] = None


t1, t2, t3 = st.tabs(["1. DPR Formatting", "2. D Form", "3. Session Count"])

with t1:
    st.subheader("DPR Formatting")
    uploaded = st.file_uploader("Upload original Excel file", type=["xlsx", "xlsm"], key="original_upload")
    if uploaded is not None and st.button("Process DPR Formatting", type="primary"):
        try:
            data = uploaded.getvalue()
            keep_vba = uploaded.name.lower().endswith(".xlsm")
            wb = load_workbook(BytesIO(data), data_only=False, keep_vba=keep_vba)
            wb = process_formatting(wb)
            st.session_state.original_bytes = data
            st.session_state.format_bytes = wb_bytes(wb)
            st.session_state.dform_bytes = None
            st.session_state.final_bytes = None
            st.success("DPR Formatting completed. Date = Left/Middle; TIME = Right/Middle; Sessions is blank.")
        except Exception as e:
            st.error(f"DPR Formatting error: {e}")
            st.exception(e)

    if st.session_state.format_bytes:
        st.download_button("Download DPR Formatting file", st.session_state.format_bytes,
                           "DPR_Formatting_Processed.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with t2:
    st.subheader("D Form")
    if st.session_state.format_bytes:
        st.info("Automatic handoff: this tab uses the latest DPR Formatting output.")
        if st.button("Process D Form", type="primary"):
            try:
                wb = load_workbook(BytesIO(st.session_state.format_bytes), data_only=False)
                wb = process_dform(wb)
                st.session_state.dform_bytes = wb_bytes(wb)
                st.session_state.final_bytes = None
                st.success("D Form completed: headings updated, two new columns added, gap-free final columns created, and calculations verified.")
            except Exception as e:
                st.error(f"D Form error: {e}")
                st.exception(e)
        if st.session_state.dform_bytes:
            st.download_button("Download D Form file", st.session_state.dform_bytes,
                               "D_Form_Processed.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("First complete Prompt 1 in DPR Formatting.")

with t3:
    st.subheader("Session Count")
    if st.session_state.dform_bytes:
        st.info("Automatic handoff: this tab uses the latest D Form output.")
        if st.button("Process Session Count", type="primary"):
            try:
                wb = load_workbook(BytesIO(st.session_state.dform_bytes), data_only=False)
                wb, total_dates, total_sessions = process_sessions(wb)
                st.session_state.final_bytes = wb_bytes(wb)
                st.success(f"Session Count completed and verified: {total_dates} unique dates, {total_sessions} total sessions. Session Summary has borders on every cell.")
            except Exception as e:
                st.error(f"Session Count error: {e}")
                st.exception(e)
        if st.session_state.final_bytes:
            st.download_button("Download Final Excel", st.session_state.final_bytes,
                               "Final_DPR_DForm_Session.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("First complete Prompt 2 in D Form.")
