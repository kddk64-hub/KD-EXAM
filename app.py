import re
import hashlib
from copy import copy
from datetime import date, datetime, timedelta
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

st.set_page_config(page_title="DPR → D Form → Session Count", layout="wide")
st.title("Excel Processing App")
st.caption("Version 5.11 — Date-wise Report, Dashboard/Summary and Excel/PDF Print Report")

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
            try:
                ws.unmerge_cells(str(rng))
            except KeyError:
                # openpyxl can retain a stale merged-cell entry after insert_cols;
                # remove the range metadata so the newly-created column can be rebuilt.
                try:
                    ws.merged_cells.ranges.remove(rng)
                except ValueError:
                    pass


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

    # Remove exactly two blank spacer columns immediately after Stu. No.
    # This matches the source layout shown in the user's workbook.
    student_col = student_col_for(ws, header_row)
    if student_col + 2 <= ws.max_column:
        if (norm_header(ws.cell(header_row, student_col + 1).value) == "" and
                norm_header(ws.cell(header_row, student_col + 2).value) == ""):
            ws.delete_cols(student_col + 1, 2)

    cols = ensure_target_columns(ws, header_row)

    # Rename the four calculated columns at the DPR stage too, so the
    # downloaded file already has the requested final headings.
    ws.cell(header_row, cols["Total"]).value = "Total Students"
    ws.cell(header_row, cols["Block"]).value = "No. of Blocks"
    ws.cell(header_row, cols["Super"]).value = "Supervisors"
    ws.cell(header_row, cols["Squad"]).value = "Int. Squad"

    # Put the two supervisor columns immediately after Int. Squad and before Sessions.
    existing_int = find_col(ws, {"int. sr. supervisor"}, header_row)
    existing_ext = find_col(ws, {"ext. sr. supervisor"}, header_row)
    if not existing_int and not existing_ext:
        squad_col = find_col(ws, {"int. squad"}, header_row)
        sessions_col = find_col(ws, {"sessions"}, header_row)
        if squad_col and sessions_col:
            ws.insert_cols(squad_col + 1, amount=2)
            sessions_col += 2
            date_col = find_col(ws, {"date"}, header_row) or date_col
            copy_date_format_and_merges(ws, date_col, squad_col + 1, header_row, 1)
            copy_date_format_and_merges(ws, date_col, squad_col + 2, header_row, 2)
            for c, value in ((squad_col + 1, 1), (squad_col + 2, 2)):
                for r in range(header_row + 1, ws.max_row + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell, MergedCell):
                        continue
                    cell.value = value
                    cell.number_format = "General"
            ws.cell(header_row, squad_col + 1).value = "Int. Sr. Supervisor"
            ws.cell(header_row, squad_col + 2).value = "Ext. Sr. Supervisor"
            ws.cell(header_row, sessions_col).value = "Sessions"
            for r in range(header_row + 1, ws.max_row + 1):
                cell = ws.cell(r, sessions_col)
                if not isinstance(cell, MergedCell):
                    cell.value = None

    # Rebuild the column map after any insertion.
    cols = {}
    for name in ("Total Students", "No. of Blocks", "Supervisors", "Int. Squad", "Int. Sr. Supervisor", "Ext. Sr. Supervisor", "Sessions"):
        c = find_col(ws, {name}, header_row)
        if c:
            cols[name] = c

    # Final requested alignment for original Date/Time columns.
    set_alignment(ws, date_col, "left", "center")
    set_alignment(ws, time_col, "right", "center")

    # Final numeric columns are centered, bold, and slightly larger.
    for h in ("Total Students", "No. of Blocks", "Supervisors", "Int. Squad", "Int. Sr. Supervisor", "Ext. Sr. Supervisor"):
        c = cols.get(h)
        if not c:
            continue
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            f.sz = max((f.sz or 11) + 1, 12)
            cell.font = f

    sessions_c = cols.get("Sessions")
    if sessions_c:
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, sessions_c)
            if not isinstance(cell, MergedCell):
                cell.value = None
                cell.alignment = Alignment(horizontal="center", vertical="center")

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
    # Destination columns are newly inserted/blank; do not unmerge stale ranges
    # left behind by openpyxl after column insertion.
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
    header_row = (find_header_row(ws, {"total students", "no. of blocks", "supervisors", "int. squad"})
                  or find_header_row(ws, {"total", "block", "super", "squad"})
                  or 1)
    student_col = student_col_for(ws, header_row)

    # Remove exactly the two blank spacer columns immediately after Stu. No.
    # Do this only when they are actually blank, so real data columns are never deleted.
    if student_col + 2 <= ws.max_column:
        gap1 = ws.cell(header_row, student_col + 1).value
        gap2 = ws.cell(header_row, student_col + 2).value
        if norm_header(gap1) == "" and norm_header(gap2) == "":
            ws.delete_cols(student_col + 1, 2)

    # Find the calculation columns, accepting the old headings as well.
    def find_any_col(*names):
        wanted = {norm_header(x) for x in names}
        for c in range(1, ws.max_column + 1):
            if norm_header(ws.cell(header_row, c).value) in wanted:
                return c
        return None

    total_col = find_any_col("Total", "Total Students")
    block_col = find_any_col("Block", "Blocks", "Blocks--", "No. of Blocks")
    super_col = find_any_col("Super", "Super--", "Supervisors")
    squad_col = find_any_col("Squad", "Squad--", "Int. Squad")
    sessions_col = find_any_col("Sessions")
    if not all([total_col, block_col, super_col, squad_col, sessions_col]):
        raise ValueError("Required Total/Block/Super/Squad/Sessions columns were not found.")

    # Calculate Total / Block / Super / Squad.
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

    # Exact requested headings.
    ws.cell(header_row, total_col).value = "Total Students"
    ws.cell(header_row, block_col).value = "No. of Blocks"
    ws.cell(header_row, super_col).value = "Supervisors"
    ws.cell(header_row, squad_col).value = "Int. Squad"

    # Remove any previously-created supervisor columns if this D Form step is run
    # repeatedly on an already-processed workbook. Keep the Sessions column.
    old_supervisor_cols = []
    for c in range(1, ws.max_column + 1):
        h = norm_header(ws.cell(header_row, c).value)
        if h in {
            "int. sr. supervisor", "int. sr. superior",
            "ext. sr. supervisor", "ext. sr. superior"
        }:
            old_supervisor_cols.append(c)
    for c in reversed(old_supervisor_cols):
        if c != sessions_col:
            ws.delete_cols(c, 1)
            if c < sessions_col:
                sessions_col -= 1

    # Refresh Squad/Sessions positions after any cleanup.
    squad_col = find_any_col("Int. Squad", "Squad", "Squad--")
    sessions_col = find_any_col("Sessions")
    if not squad_col or not sessions_col:
        raise ValueError("Int. Squad or Sessions column was lost while preparing D Form.")

    # Insert the two requested columns immediately after Int. Squad.
    int_sr_col = squad_col + 1
    ws.insert_cols(int_sr_col, amount=2)
    ext_sr_col = int_sr_col + 1

    # Copy Date-column formatting and vertical merge structure to both columns.
    # This also clears the old date values in the new columns.
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column is required for supervisor columns.")
    copy_date_format_and_merges(ws, date_col, int_sr_col, header_row, 1)
    copy_date_format_and_merges(ws, date_col, ext_sr_col, header_row, 2)

    # The Date column's number format would otherwise display 1/2 as dates
    # (for example 01/01/1900 and 02/01/1900). Keep the Date-style layout,
    # but make these supervisor values real numeric 1/2 with General format.
    for c, value in ((int_sr_col, 1), (ext_sr_col, 2)):
        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = value
            cell.number_format = "General"

    # Sessions must use the same formatting/merge structure as Ext. Sr. Supervisor,
    # while remaining completely blank. Rebuild the Sessions column from Ext. Sr. Supervisor.
    sessions_col = find_any_col("Sessions")
    unmerge_ranges_in_column(ws, sessions_col)
    copy_col_dimension(ws, ext_sr_col, sessions_col)
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, ext_sr_col)
        dst = ws.cell(r, sessions_col)
        if not isinstance(dst, MergedCell):
            copy_cell_style(src, dst)
            dst.value = None
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == ext_sr_col and rng.max_col == ext_sr_col:
            ws.merge_cells(start_row=rng.min_row, start_column=sessions_col,
                           end_row=rng.max_row, end_column=sessions_col)
            top = ws.cell(rng.min_row, sessions_col)
            top.value = None
            copy_cell_style(ws.cell(rng.min_row, ext_sr_col), top)

    ws.cell(header_row, int_sr_col).value = "Int. Sr. Supervisor"
    ws.cell(header_row, ext_sr_col).value = "Ext. Sr. Supervisor"
    ws.cell(header_row, sessions_col).value = "Sessions"

    # Center, bold, and slightly enlarge all numeric values from Total Students
    # through Ext. Sr. Supervisor.
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
    # The Summary "Dates" column must be left aligned.
    if ws.max_column >= 2:
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, 2)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def extract_workbook_dates(ws, header_row=1):
    """Return sorted unique actual dates present in the main sheet."""
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        return []
    out = set()
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        k = date_key(v)
        if k:
            try:
                out.add(datetime.strptime(k, "%Y-%m-%d").date())
            except ValueError:
                pass
    return sorted(out)


def add_date_range_report(wb, start_date, end_date):
    """Filter the main worksheet in-place to the selected inclusive date range.
    Preserves cell values, styles, dimensions and applicable merged ranges.
    """
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column सापडला नाही.")

    row_dates = {}
    current = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v not in (None, ""):
            k = date_key(v)
            try:
                current = datetime.strptime(k, "%Y-%m-%d").date()
            except (ValueError, TypeError):
                current = None
        row_dates[r] = current

    original_merges = [str(rng) for rng in ws.merged_cells.ranges]
    keep_rows = {
        r for r in range(header_row + 1, ws.max_row + 1)
        if row_dates.get(r) is not None and start_date <= row_dates[r] <= end_date
    }

    # Unmerge before deleting rows so merged-cell placeholders do not block deletion.
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    old_max_row = ws.max_row
    deleted_before = 0
    row_map = {}
    for r in range(1, old_max_row + 1):
        if r <= header_row or r in keep_rows:
            row_map[r] = r - deleted_before
        else:
            deleted_before += 1

    for r in range(old_max_row, header_row, -1):
        if r not in keep_rows:
            ws.delete_rows(r, 1)

    # Recreate merged ranges only when the entire original range survived filtering.
    for merge_ref in original_merges:
        min_col, min_row, max_col, max_row = range_boundaries(merge_ref)
        if all(rr <= header_row or rr in keep_rows for rr in range(min_row, max_row + 1)):
            new_min_row = row_map[min_row]
            new_max_row = row_map[max_row]
            try:
                ws.merge_cells(start_row=new_min_row, start_column=min_col,
                               end_row=new_max_row, end_column=max_col)
            except Exception:
                pass

    return ws

def build_date_range_report(wb, start_date, end_date):
    """Return a workbook whose main report sheet contains only the selected date range.
    If either bound is None, the full available date range is used.
    """
    src = get_main_sheet(wb)
    original_title = src.title
    all_dates = extract_workbook_dates(src, find_header_row(src, {"date", "time"}) or 1)
    if not all_dates:
        raise ValueError("Date column मध्ये कोणतीही तारीख सापडली नाही.")
    if start_date is None:
        start_date = all_dates[0]
    if end_date is None:
        end_date = all_dates[-1]
    if start_date > end_date:
        raise ValueError("From Date हा To Date पेक्षा मोठा असू शकत नाही.")

    # Keep a copy of session counts before replacing the main sheet.
    session_counts = {}
    header_row = find_header_row(src, {"date", "time"}) or 1
    date_col = find_col(src, {"date"}, header_row)
    sessions_col = find_col(src, {"sessions"}, header_row)
    current_date = None
    for r in range(header_row + 1, src.max_row + 1):
        dv = src.cell(r, date_col).value
        if dv not in (None, ""):
            k = date_key(dv)
            try:
                current_date = datetime.strptime(k, "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if current_date and sessions_col:
            sv = src.cell(r, sessions_col).value
            if sv not in (None, ""):
                try:
                    session_counts[current_date] = int(sv)
                except Exception:
                    pass

    filtered = add_date_range_report(wb, start_date, end_date)
    # Keep the main worksheet visible; filtering it in-place avoids an empty-workbook error.
    filtered.title = original_title

    # For Session Count reports, also filter Session Summary to the same date range.
    if "Session Summary" in wb.sheetnames and session_counts:
        del wb["Session Summary"]
        summary = wb.create_sheet("Session Summary")
        headers = ["Session Name", "Dates", "Total Dates", "Actual Sessions", "Rs."]
        summary.append(headers)
        for c in range(1, 6):
            summary.cell(1, c).font = Font(bold=True)
            summary.cell(1, c).alignment = Alignment(horizontal="center", vertical="center")
        cats = {i: [] for i in range(1, 8)}
        selected_counts = {d:n for d,n in session_counts.items() if start_date <= d <= end_date}
        for d, n in selected_counts.items():
            if 1 <= n <= 7:
                cats[n].append(d.strftime("%Y-%m-%d"))
        for n in range(1, 8):
            ds = sorted(cats[n])
            total = len(ds)
            actual = n * total
            summary.append([n, ", ".join(display_date(d) for d in ds), total, actual, actual * 105])
        total_dates = len(selected_counts)
        total_sessions = sum(selected_counts.values())
        summary.append(["TOTAL", "", total_dates, total_sessions, total_sessions * 105])
        summary.column_dimensions["A"].width = 15
        summary.column_dimensions["B"].width = 55
        summary.column_dimensions["C"].width = 15
        summary.column_dimensions["D"].width = 18
        summary.column_dimensions["E"].width = 15
        apply_summary_borders(summary)

    return wb


def resolve_date_range(start_choice, end_choice, available_dates):
    """Convert dropdown selections to an actual date range; ALL means full range."""
    if not available_dates:
        raise ValueError("Date column मध्ये कोणतीही तारीख सापडली नाही.")
    if start_choice == "ALL" or end_choice == "ALL":
        return available_dates[0], available_dates[-1]
    return start_choice, end_choice


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
        # Fill the Sessions column with the calculated session count for each date.
        # The Sessions column has the same vertical merge structure as Ext. Sr. Supervisor,
        # so the value is written to the visible/top-left cell of each merged date group.
        current_date = ""
        date_for_row = {}
        for r in range(header_row + 1, ws.max_row + 1):
            d = ws.cell(r, date_col).value
            if d is not None and str(d).strip() != "":
                current_date = date_key(d)
            date_for_row[r] = current_date

        for r in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(r, sessions_col)
            if isinstance(cell, MergedCell):
                continue
            d = date_for_row.get(r, "")
            cell.value = counts.get(d, None)
            cell.number_format = "General"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            f = copy(cell.font)
            f.bold = True
            f.sz = max((f.sz or 11) + 1, 12)
            cell.font = f

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
    for r in range(2, summary.max_row + 1):
        summary.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    apply_summary_borders(summary)

    return wb, total_dates, total_sessions



def report_source_bytes():
    """Return the full processed workbook, never a date/class-filtered report copy."""
    return (
        st.session_state.get("session_full_bytes")
        or st.session_state.get("dform_bytes")
    )


def workbook_rows_for_range(wb, start_date, end_date):
    """Extract the main sheet rows for an inclusive date range, forward-filling merged dates."""
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    if not date_col:
        raise ValueError("Date column सापडला नाही.")
    headers = [ws.cell(header_row, c).value or f"Column {c}" for c in range(1, ws.max_column + 1)]
    rows = []
    current_date = None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if v not in (None, ""):
            k = date_key(v)
            try:
                current_date = datetime.strptime(k, "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if current_date is None or not (start_date <= current_date <= end_date):
            continue
        vals = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            vals.append(None if isinstance(cell, MergedCell) else cell.value)
        vals[date_col - 1] = current_date
        rows.append(vals)
    return headers, rows


def filtered_workbook_copy(source_bytes, start_date, end_date):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    build_date_range_report(wb, start_date, end_date)
    return wb


def make_pdf_report(source_bytes, start_date, end_date, title="Print Report"):
    """Create a readable PDF using a custom wide landscape page instead of A4.
    The page width expands with the number of columns so text is not forced into
    tiny A4-width cells.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    headers, rows = workbook_rows_for_range(wb, start_date, end_date)
    out = BytesIO()

    # Custom wide landscape page: no A4-width restriction.
    # Give each column enough room for readable text, while allowing very wide
    # reports to span multiple physical pages vertically.
    ncols = max(1, len(headers))
    col_width = 24 * mm
    page_width = max(320 * mm, min(650 * mm, ncols * col_width + 20 * mm))
    page_height = 210 * mm

    doc = SimpleDocTemplate(
        out,
        pagesize=(page_width, page_height),
        rightMargin=8 * mm,
        leftMargin=8 * mm,
        topMargin=8 * mm,
        bottomMargin=8 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PrintTitle", parent=styles["Title"], fontSize=15, leading=18
    )
    normal_style = ParagraphStyle(
        "PrintNormal", parent=styles["Normal"], fontSize=9, leading=11
    )
    cell_style = ParagraphStyle(
        "PrintCell", parent=styles["Normal"], fontSize=8, leading=9, wordWrap="CJK"
    )
    header_style = ParagraphStyle(
        "PrintHeader", parent=cell_style, fontName="Helvetica-Bold"
    )

    elements = [
        Paragraph(title, title_style),
        Paragraph(
            f"Date Range: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}",
            normal_style,
        ),
        Spacer(1, 4 * mm),
    ]

    if not rows:
        elements.append(Paragraph("No data found for the selected date range.", normal_style))
    else:
        def cell_text(v):
            if isinstance(v, datetime):
                return v.strftime("%d/%m/%Y %I:%M %p")
            if isinstance(v, date):
                return v.strftime("%d/%m/%Y")
            return "" if v is None else str(v)

        data = [[Paragraph(cell_text(h), header_style) for h in headers]]
        for row in rows:
            data.append([Paragraph(cell_text(v), cell_style) for v in row])

        available_width = page_width - 16 * mm
        widths = [available_width / ncols] * ncols
        table = Table(data, repeatRows=1, colWidths=widths, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        elements.append(table)

    doc.build(elements)
    return out.getvalue()


def find_class_col(ws, header_row):
    return find_col(ws, {
        "class", "class name", "class_name", "standard", "std", "std.",
        "standard name", "class/standard", "class/div", "class / div",
        "class division", "इयत्ता"
    }, header_row)


def extract_workbook_classes(ws, header_row=1):
    class_col = find_class_col(ws, header_row)
    if not class_col:
        return []
    values, current = set(), None
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, class_col).value
        if v not in (None, ""):
            current = str(v).strip()
        if current:
            values.add(current)
    return sorted(values, key=str.lower)


def available_report_classes(source_bytes):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    return extract_workbook_classes(ws, header_row)


def filter_workbook_by_date_and_class(wb, start_date, end_date, selected_classes=None):
    ws = get_main_sheet(wb)
    header_row = find_header_row(ws, {"date", "time"}) or 1
    date_col = find_col(ws, {"date"}, header_row)
    class_col = find_class_col(ws, header_row)

    if not date_col:
        raise ValueError("Date column सापडला नाही.")
    if selected_classes and not class_col:
        raise ValueError("Excel मध्ये Class/Standard column सापडला नाही.")

    selected = {str(x).strip() for x in (selected_classes or [])}
    row_info, current_date, current_class = {}, None, None

    for r in range(header_row + 1, ws.max_row + 1):
        dv = ws.cell(r, date_col).value
        if dv not in (None, ""):
            try:
                current_date = datetime.strptime(date_key(dv), "%Y-%m-%d").date()
            except Exception:
                current_date = None
        if class_col:
            cv = ws.cell(r, class_col).value
            if cv not in (None, ""):
                current_class = str(cv).strip()
        row_info[r] = (current_date, current_class)

    keep_rows = {
        r for r, (d, c) in row_info.items()
        if d is not None and start_date <= d <= end_date
        and (not selected or c in selected)
    }

    original_merges = [str(rng) for rng in ws.merged_cells.ranges]
    old_max_row = ws.max_row
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    row_map, deleted_before = {}, 0
    for r in range(1, old_max_row + 1):
        if r <= header_row or r in keep_rows:
            row_map[r] = r - deleted_before
        else:
            deleted_before += 1

    for r in range(old_max_row, header_row, -1):
        if r not in keep_rows:
            ws.delete_rows(r, 1)

    for ref in original_merges:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
        if all(rr <= header_row or rr in keep_rows for rr in range(min_row, max_row + 1)):
            try:
                ws.merge_cells(
                    start_row=row_map[min_row], start_column=min_col,
                    end_row=row_map[max_row], end_column=max_col
                )
            except Exception:
                pass
    return wb


def date_class_range_controls(source_bytes, prefix):
    dates = available_report_dates(source_bytes) if source_bytes else []
    classes = available_report_classes(source_bytes) if source_bytes else []
    if not dates:
        return None, None, []

    choices = ["ALL"] + dates
    c1, c2 = st.columns(2)
    with c1:
        start = st.selectbox(
            "From Date", choices, index=0,
            format_func=lambda x: "ALL" if x == "ALL" else x.strftime("%d/%m/%Y"),
            key=f"{prefix}_from",
        )
    with c2:
        valid = choices if start == "ALL" else ["ALL"] + [d for d in dates if d >= start]
        old = st.session_state.get(f"{prefix}_to")
        idx = valid.index(old) if old in valid else 0
        end = st.selectbox(
            "To Date", valid, index=idx,
            format_func=lambda x: "ALL" if x == "ALL" else x.strftime("%d/%m/%Y"),
            key=f"{prefix}_to",
        )

    # Class filter: keep the page compact and show classes only inside a dropdown/popover.
    # Each class is selectable with its own checkbox. Existing behaviour is preserved: all
    # available classes are selected by default.
    state_key = f"{prefix}_classes"
    if state_key not in st.session_state:
        st.session_state[state_key] = list(classes)
    else:
        # Keep only currently available classes if the source workbook changes.
        st.session_state[state_key] = [c for c in st.session_state[state_key] if c in classes]

    selected = st.session_state[state_key]
    selected_label = "All Classes" if len(selected) == len(classes) else (
        f"{len(selected)} Class(es) selected" if selected else "No Class selected"
    )

    with st.popover(f"Class: {selected_label}", use_container_width=True):
        st.markdown("**Select Class**")

        select_all = st.checkbox(
            "Select All",
            value=(len(selected) == len(classes)),
            key=f"{prefix}_class_select_all",
        )

        if select_all and len(selected) != len(classes):
            selected = list(classes)
        elif not select_all and len(selected) == len(classes):
            selected = []

        updated = []
        for i, cls in enumerate(classes):
            checked = cls in selected
            if st.checkbox(str(cls), value=checked, key=f"{prefix}_class_{i}"):
                updated.append(cls)

        # The individual checkboxes are the final source of truth.
        st.session_state[state_key] = updated
        selected = updated

    return resolve_date_range(start, end, dates) + (selected,)


def filtered_report_bytes(source, start_date, end_date, selected_classes):
    wb = load_workbook(BytesIO(source), data_only=False)
    filter_workbook_by_date_and_class(wb, start_date, end_date, selected_classes)
    return wb_bytes(wb)


def available_report_dates(source_bytes):
    wb = load_workbook(BytesIO(source_bytes), data_only=False)
    ws = get_main_sheet(wb)
    return extract_workbook_dates(ws, find_header_row(ws, {"date", "time"}) or 1)


for key in ("original_bytes", "format_bytes", "dform_bytes", "final_bytes", "range_report_bytes", "dform_range_report_bytes", "report_range", "dform_report_range", "datewise_excel_bytes", "print_excel_bytes", "print_pdf_bytes", "session_full_bytes", "session_report_bytes", "t1_report_excel_bytes", "t1_report_pdf_bytes"):
    if key not in st.session_state:
        st.session_state[key] = None


def date_range_controls(source_bytes, prefix):
    dates = available_report_dates(source_bytes) if source_bytes else []
    if not dates:
        return None, None
    choices = ["ALL"] + dates
    c1, c2 = st.columns(2)
    with c1:
        start = st.selectbox("From Date", choices, index=0,
                             format_func=lambda x: "ALL" if x == "ALL" else x.strftime("%d/%m/%Y"),
                             key=f"{prefix}_from")
    with c2:
        valid = choices if start == "ALL" else ["ALL"] + [d for d in dates if d >= start]
        old = st.session_state.get(f"{prefix}_to")
        idx = valid.index(old) if old in valid else 0
        end = st.selectbox("To Date", valid, index=idx,
                           format_func=lambda x: "ALL" if x == "ALL" else x.strftime("%d/%m/%Y"),
                           key=f"{prefix}_to")
    return resolve_date_range(start, end, dates)


def render_print_report(source, start_date, end_date, selected_classes, prefix, title="DPR / D Form Print Report"):
    """Direct print/download actions: no separate Generate step."""
    st.markdown("---")
    st.subheader("🖨️ Print Report")
    st.caption(
        f"Print कालावधी: {start_date.strftime('%d/%m/%Y')} ते "
        f"{end_date.strftime('%d/%m/%Y')}"
    )

    try:
        filtered_bytes = filtered_report_bytes(
            source, start_date, end_date, selected_classes
        )
        print_excel = filtered_bytes
        print_pdf = make_pdf_report(
            filtered_bytes, start_date, end_date, title
        )

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "🖨️ Print Excel",
                print_excel,
                f"{prefix}_Print_Report.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{prefix}_direct_print_excel",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "🖨️ Print PDF",
                print_pdf,
                f"{prefix}_Print_Report.pdf",
                "application/pdf",
                key=f"{prefix}_direct_print_pdf",
                use_container_width=True,
            )
    except Exception as e:
        st.error(f"Print Report error: {e}")
        st.exception(e)


t1, t2, t3, t4 = st.tabs([
    "1. DPR + D Form", "2. Session Count", "3. Date-wise Report", "4. Dashboard / Summary"
])

with t1:
    st.subheader("DPR Formatting + D Form")
    st.caption("Upload → full Date/Time formatting + merge + D Form automatically → then choose Date Range + Class.")
    uploaded = st.file_uploader(
        "Upload original Excel file",
        type=["xlsx", "xlsm"],
        key="original_upload",
    )

    if uploaded is not None:
        try:
            data = uploaded.getvalue()
            keep_vba = uploaded.name.lower().endswith(".xlsm")
            upload_signature = hashlib.sha256(data).hexdigest()

            # IMPORTANT: format the COMPLETE workbook first. Date/Class filtering is
            # deliberately NOT done here, so every later tab receives the full
            # formatted workbook with the original Date/Time merge structure intact.
            if st.session_state.get("processed_upload_signature") != upload_signature:
                wb = load_workbook(BytesIO(data), data_only=False, keep_vba=keep_vba)
                wb = process_formatting(wb)
                formatted_bytes = wb_bytes(wb)

                wb = load_workbook(BytesIO(formatted_bytes), data_only=False, keep_vba=keep_vba)
                wb = process_dform(wb)
                full_dform_bytes = wb_bytes(wb)

                st.session_state.original_bytes = data
                st.session_state.format_bytes = formatted_bytes
                st.session_state.dform_bytes = full_dform_bytes
                st.session_state.final_bytes = None
                st.session_state.session_full_bytes = None
                st.session_state.session_report_bytes = None
                st.session_state.datewise_excel_bytes = None
                st.session_state.print_excel_bytes = None
                st.session_state.print_pdf_bytes = None
                st.session_state.processed_upload_signature = upload_signature

                # Reset report outputs when a new workbook is uploaded.
                for k in list(st.session_state.keys()):
                    if k.startswith(("t1_", "t2_", "t3_", "t4_")) and k.endswith(("_bytes", "_from", "_to", "_classes")):
                        if k.endswith("_bytes"):
                            st.session_state[k] = None

                st.success("पूर्ण Excel Date/Time formatting, merge आणि D Form processing पूर्ण झाले. आता Date Range + Class निवडा.")

            source = st.session_state.dform_bytes
            dates = available_report_dates(source)
            classes = available_report_classes(source)

            start_date, end_date, selected_classes = date_class_range_controls(source, "dpr_report")
            st.caption(
                f"Report कालावधी: {start_date.strftime('%d/%m/%Y')} ते "
                f"{end_date.strftime('%d/%m/%Y')}"
            )

            st.download_button(
                "Download Full Formatted D Form Excel",
                source,
                "DPR_DForm_Processed_FULL.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="t1_download_full_dform",
            )

            if st.button("Generate Tab 1 Date/Class Report", type="primary", key="t1_generate_report"):
                try:
                    st.session_state.t1_report_excel_bytes = filtered_report_bytes(
                        source, start_date, end_date, selected_classes
                    )
                    st.session_state.t1_report_pdf_bytes = None
                    st.success("Tab 1 Date/Class report तयार झाला.")
                except Exception as e:
                    st.error(f"Tab 1 report error: {e}")
                    st.exception(e)

            if st.session_state.get("t1_report_excel_bytes"):
                st.download_button(
                    "Download Tab 1 Excel",
                    st.session_state.t1_report_excel_bytes,
                    "DPR_DForm_Date_Class_Report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t1_download_report_excel",
                )
                # Print section is based on the selected Tab-1 range/class, but the
                # source remains the FULL formatted workbook.
                render_print_report(
                    source, start_date, end_date, selected_classes,
                    "t1", "DPR / D Form Print Report"
                )
        except Exception as e:
            st.error(f"DPR + D Form error: {e}")
            st.exception(e)
    else:
        st.info("Excel upload करा. Upload होताच पूर्ण workbook चे Date/Time formatting आणि D Form processing आपोआप होईल.")


with t2:
    st.subheader("Session Count + Date Range + Class")
    # Always start from the FULL formatted workbook produced in Tab 1.
    source = st.session_state.get("dform_bytes")
    if source:
        start_date, end_date, selected_classes = date_class_range_controls(source, "session_report")
        st.caption(
            f"Session Report कालावधी: {start_date.strftime('%d/%m/%Y')} ते "
            f"{end_date.strftime('%d/%m/%Y')}"
        )

        if st.button("Process Session Count", type="primary", key="t2_process_session"):
            try:
                # 1) Process the FULL workbook and save it as the common downstream source.
                full_wb = load_workbook(BytesIO(source), data_only=False)
                full_wb, total_dates, total_sessions = process_sessions(full_wb)
                st.session_state.session_full_bytes = wb_bytes(full_wb)
                st.session_state.final_bytes = st.session_state.session_full_bytes

                # 2) Separately create only this tab's requested report. This copy is
                # filtered AFTER formatting, so Date/Time merges/styles are retained.
                report_wb = load_workbook(BytesIO(source), data_only=False)
                filter_workbook_by_date_and_class(
                    report_wb, start_date, end_date, selected_classes
                )
                report_wb, report_dates, report_sessions = process_sessions(report_wb)
                st.session_state.session_report_bytes = wb_bytes(report_wb)

                st.success(
                    f"Session Count completed: full workbook {total_dates} unique dates, "
                    f"{total_sessions} total sessions. Selected report: "
                    f"{report_dates} dates, {report_sessions} sessions."
                )
            except Exception as e:
                st.error(f"Session Count error: {e}")
                st.exception(e)

        if st.session_state.get("session_report_bytes"):
            st.download_button(
                "Download Session Count Excel",
                st.session_state.session_report_bytes,
                "Final_DPR_DForm_Session_Selected.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="t2_download_report",
            )
            render_print_report(
                source, start_date, end_date, selected_classes,
                "t2", "Session Count Print Report"
            )
    else:
        st.warning("प्रथम Tab 1 मध्ये Excel upload करा.")


with t3:
    st.subheader("Date-wise Report + Date Range + Class")
    # Prefer the FULL Session Count workbook if Tab 2 has been processed; otherwise
    # use the FULL D Form workbook from Tab 1. Never use a filtered report as source.
    source = report_source_bytes()
    if source:
        start_date, end_date, selected_classes = date_class_range_controls(source, "datewise_report")
        st.caption(
            f"Date-wise Report: {start_date.strftime('%d/%m/%Y')} ते "
            f"{end_date.strftime('%d/%m/%Y')}"
        )
        try:
            wb = load_workbook(BytesIO(source), data_only=False)
            filter_workbook_by_date_and_class(wb, start_date, end_date, selected_classes)
            headers, rows = workbook_rows_for_range(wb, start_date, end_date)
            st.write(f"एकूण records: **{len(rows)}**")
            if rows:
                display_rows = []
                for row in rows:
                    item = {}
                    for i, h in enumerate(headers):
                        key = str(h).strip() or f"Column {i+1}"
                        v = row[i]
                        if isinstance(v, (datetime, date)):
                            v = v.strftime("%d/%m/%Y") if isinstance(v, date) else v.strftime("%d/%m/%Y %I:%M %p")
                        item[key] = v
                    display_rows.append(item)
                st.dataframe(display_rows, use_container_width=True, hide_index=True)

            if st.button("Generate Date-wise Excel Report", type="primary", key="t3_generate_datewise"):
                st.session_state.datewise_excel_bytes = filtered_report_bytes(
                    source, start_date, end_date, selected_classes
                )
                st.success("Date-wise Excel Report तयार झाला.")

            if st.session_state.get("datewise_excel_bytes"):
                st.download_button(
                    "Download Date-wise Excel",
                    st.session_state.datewise_excel_bytes,
                    "Date_Wise_Report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t3_download_datewise",
                )
                render_print_report(
                    source, start_date, end_date, selected_classes,
                    "t3", "Date-wise Print Report"
                )
        except Exception as e:
            st.error(f"Date-wise Report error: {e}")
            st.exception(e)
    else:
        st.warning("प्रथम Tab 1 मध्ये Excel upload करा.")


with t4:
    st.subheader("Dashboard / Summary Report")
    source = report_source_bytes()
    if source:
        try:
            today = date.today()
            wb = load_workbook(BytesIO(source), data_only=False)
            headers, rows = workbook_rows_for_range(wb, today, today)
            st.markdown(f"### आजची तारीख: {today.strftime('%d/%m/%Y')}")
            st.metric("आजचे Records", len(rows))
            norm_to_idx = {norm_header(h): i for i, h in enumerate(headers)}
            for label, aliases in [
                ("Total Students", ["total students", "total"]),
                ("No. of Blocks", ["no. of blocks", "block", "blocks"]),
                ("Supervisors", ["supervisors", "super"]),
                ("Int. Squad", ["int. squad", "squad"]),
                ("Sessions", ["sessions"]),
            ]:
                idx = next((norm_to_idx[a] for a in aliases if a in norm_to_idx), None)
                if idx is not None:
                    nums = []
                    for row in rows:
                        try:
                            if row[idx] not in (None, ""):
                                nums.append(float(row[idx]))
                        except Exception:
                            pass
                    if nums:
                        value = sum(nums)
                        st.write(f"**{label}:** {int(value) if value.is_integer() else value}")
            if rows:
                st.dataframe(
                    [{str(h): row[i] for i, h in enumerate(headers)} for row in rows],
                    use_container_width=True,
                    hide_index=True,
                )
                # Summary print uses today's records from the same FULL formatted source.
                render_print_report(
                    source, today, today, [], "t4", "Dashboard / Summary Print Report"
                )
            else:
                st.info("आजच्या current date साठी Excel मध्ये data उपलब्ध नाही.")
        except Exception as e:
            st.error(f"Dashboard error: {e}")
            st.exception(e)
    else:
        st.warning("प्रथम Tab 1 मध्ये Excel upload करा.")
