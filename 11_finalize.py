"""Step 11: Final cleanups — enclosure labeling and column trimming.

Replaces two separate scripts:
  - 11_final.py   (fill missing enclosure/tray labels, insert MST/DEMUX rows)
  - 12_cleanup_splitter_and_trim.py  (shift splitter metadata, trim columns)

Reads:  output/Combined_Final_Shifted_B3_Labeled.xlsx
Writes: output/Asbuilt_Workbook_post12.xlsx
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

import openpyxl
from openpyxl.styles import PatternFill

from naming_utils import (
    is_location_sheet,
    find_header_row,
    find_col_by_header,
    find_all_cols_by_header,
    safe_fill_hex,
)

INPUT_FILE  = Path("output") / "Combined_Reordered_With_OTE.xlsx"
OUTPUT_FILE = Path("output") / "Asbuilt_Workbook_post12.xlsx"

legacy_name_rx = re.compile(
    r'^(?P<prefix>[A-Z]{2}[A-Z0-9]+?)(?P<type>S|D)?(?P<num>\d{3,4})$',
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Part A  (from 11_final.py): Enclosure labels, DEMUX + MST row insertion
# ---------------------------------------------------------------------------

def _fill_enclosure_labels(sheet, sheet_name: str) -> None:
    """Populate B3 (Enclosure) and B4 (Trays) based on naming convention."""
    enc_cell   = sheet.cell(row=3, column=2)
    trays_cell = sheet.cell(row=4, column=2)

    if enc_cell.value and str(enc_cell.value).strip() and \
       trays_cell.value and str(trays_cell.value).strip():
        return  # already filled

    name_up = sheet_name.upper()
    enc_val = trays_val = None

    if "_FT_" in name_up or name_up.endswith("_FT"):
        enc_val, trays_val = "2 PORT OTE", 1
    elif "_SE_" in name_up or name_up.endswith("_SE") or re.search(r'S\d+$', name_up):
        enc_val, trays_val = "COMMSCOPE FOSC 450-D", 2
    elif re.search(r'D\d+$', name_up):
        enc_val, trays_val = "COMMSCOPE FOSC 450-B", 1
    else:
        m = legacy_name_rx.match(name_up)
        if m:
            t = (m.group("type") or "").upper()
            if t == "S":
                enc_val, trays_val = "COMMSCOPE FOSC 450-D", 2
            elif t == "D":
                enc_val, trays_val = "COMMSCOPE FOSC 450-B", 1

    if enc_val and not (enc_cell.value and str(enc_cell.value).strip()):
        enc_cell.value = enc_val
    if trays_val is not None and not (trays_cell.value and str(trays_cell.value).strip()):
        trays_cell.value = trays_val


def _insert_demux_mst_rows(sheet) -> None:
    """Insert DEMUX and MST rows when both 1x2 and 1x32 devices are present."""
    sheaths_row = None
    for cell in sheet['A']:
        if cell.value == "SHEATHS":
            sheaths_row = cell.row
            break
    if sheaths_row is None:
        return

    meta_blank_row = 8  # row after metadata block

    first_conn_row = None
    for r in range(meta_blank_row + 1, sheaths_row):
        val_a = sheet.cell(row=r, column=1).value
        val_b = sheet.cell(row=r, column=2).value
        if (val_a is None or not str(val_a).strip()) and \
           (val_b is None or not str(val_b).strip()):
            break
        if val_a and not (val_b and str(val_b).strip()):
            first_conn_row = r
            break

    if first_conn_row is None:
        return

    found_1x2 = found_1x32 = found_demux = False
    for r in range(meta_blank_row + 1, first_conn_row):
        cell_a = str(sheet.cell(row=r, column=1).value or "").upper()
        cell_b = str(sheet.cell(row=r, column=2).value or "").upper()
        if "1X32" in cell_a or "1X32" in cell_b:
            found_1x32 = True
        if "1X2" in cell_a or "1X2" in cell_b:
            found_1x2 = True
        if "DEMUX" in cell_a:
            found_demux = True

    if not (found_1x2 and found_1x32):
        return

    if not found_demux:
        sheet.insert_rows(first_conn_row)
        sheet.cell(row=first_conn_row, column=1).value = "DEMUX"
        sheet.cell(row=first_conn_row, column=2).value = "4CH"
        first_conn_row += 1
        if sheaths_row:
            sheaths_row += 1

    # Check for existing MST line
    blank_above = sheaths_row - 1 if sheaths_row else None
    has_mst = any(
        isinstance(sheet.cell(row=r, column=1).value, str) and
        sheet.cell(row=r, column=1).value.strip().upper() == "MST"
        for r in range(first_conn_row, blank_above or sheaths_row)
    )

    if not has_mst and blank_above:
        sheet.insert_rows(blank_above)
        sheet.cell(row=blank_above, column=1).value = "MST"
        sheet.cell(row=blank_above, column=4).value = "24CT"


def run_part_a(wb) -> None:
    for name in wb.sheetnames:
        sheet = wb[name]
        if sheet.cell(row=3, column=1).value != "Enclosure:" or \
           sheet.cell(row=4, column=1).value != "No. of Trays:":
            continue
        _fill_enclosure_labels(sheet, name)
        _insert_demux_mst_rows(sheet)


# ---------------------------------------------------------------------------
# Part B  (from 12_cleanup_splitter_and_trim.py): Column shifting + trimming
# ---------------------------------------------------------------------------

def _norm(v) -> str:
    return str(v).strip().upper() if v is not None else ""


def _clear_range(ws, r_min, r_max, c_min, c_max) -> None:
    empty = PatternFill()
    for r in range(r_min, r_max + 1):
        for c in range(c_min, c_max + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = empty


def _clear_optical_splitter_sheath_uuid(ws) -> bool:
    """Clear SHEATH UUID column in the OPTICAL SPLITTERS sub-table."""
    opt_row = None
    for r in range(1, min(ws.max_row, 400) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "OPTICAL SPLITTERS" in v.upper():
            opt_row = r
            break
    if opt_row is None:
        return False

    sheath_pos = None
    for r in range(opt_row, min(ws.max_row, opt_row + 40) + 1):
        for c in range(1, min(ws.max_column, 160) + 1):
            if _norm(ws.cell(r, c).value) == "SHEATH UUID":
                sheath_pos = (r, c)
                break
        if sheath_pos:
            break

    if not sheath_pos:
        return False

    hdr_row, sheath_uuid_col = sheath_pos
    last_col = sheath_uuid_col
    for c in range(sheath_uuid_col + 1, min(ws.max_column, 160) + 1):
        if ws.cell(hdr_row, c).value and str(ws.cell(hdr_row, c).value).strip():
            break
        last_col = c

    _clear_range(ws, hdr_row, ws.max_row, sheath_uuid_col, last_col)
    return True


def _process_sheet_cleanup(ws, name: str) -> Tuple[bool, bool, bool]:
    """Shift splitter metadata left, clear optical SHEATH UUID, trim right columns."""
    hdr_row = find_header_row(ws)
    if hdr_row is None:
        return (False, False, False)

    conn_col = find_col_by_header(ws, hdr_row, "CONNECTION")
    if conn_col is None:
        return (False, False, False)

    buffer_cols = find_all_cols_by_header(ws, hdr_row, "BUFFER")
    buffer_after = next((c for c in buffer_cols if c > conn_col), conn_col)
    dest_start = buffer_after + 1

    sheath_uuid_cols = find_all_cols_by_header(ws, hdr_row, "SHEATH UUID")
    trim_col = next((c for c in sheath_uuid_cols if c > conn_col), None)

    device_name_col = find_col_by_header(ws, hdr_row, "DEVICE NAME", after_col=conn_col)
    device_uuid_col = find_col_by_header(ws, hdr_row, "DEVICE UUID", after_col=conn_col)

    has_splitter = any(
        isinstance(ws.cell(r, 1).value, str) and "OPTICAL SPLITTERS" in ws.cell(r, 1).value.upper()
        for r in range(1, min(ws.max_row, 400) + 1)
    )

    shifted = False
    if has_splitter and device_name_col and device_uuid_col:
        keep = {"FIBER", "BUFFER", "END ENCLOSURE", "END ENCLOSURE ", "START ENCLOSURE",
                "SHEATH NAME", "SHEATH UUID"}
        right_cols = [c for c in range(conn_col + 1, device_uuid_col)
                      if _norm(ws.cell(hdr_row, c).value) and
                      _norm(ws.cell(hdr_row, c).value) not in keep]

        primary = next((c for c in sheath_uuid_cols if c <= conn_col), 1)
        last_row = hdr_row
        blank_run = 0
        for r in range(hdr_row + 1, ws.max_row + 1):
            v = ws.cell(r, primary).value
            if v is None or (isinstance(v, str) and not v.strip()):
                blank_run += 1
                if blank_run >= 5:
                    break
            else:
                blank_run = 0
                last_row = r

        for r in range(hdr_row + 1, last_row + 1):
            dname = ws.cell(r, device_name_col).value
            if not isinstance(dname, str):
                continue
            if "1X2" not in dname.upper() and "1X32" not in dname.upper():
                continue
            src_start = next(
                (c for c in right_cols
                 if ws.cell(r, c).value is not None and
                 (not isinstance(ws.cell(r, c).value, str) or ws.cell(r, c).value.strip())),
                None,
            )
            if src_start is None:
                continue
            src_end = device_uuid_col - 1
            for i in range(src_end - src_start + 1):
                dst_c = dest_start + i
                if trim_col and dst_c >= trim_col:
                    break
                ws.cell(r, dst_c).value = ws.cell(r, src_start + i).value
            for c in range(src_start, src_end + 1):
                ws.cell(r, c).value = None
            shifted = True

    cleared = _clear_optical_splitter_sheath_uuid(ws) if has_splitter else False

    trimmed = False
    if trim_col:
        ws.delete_cols(trim_col, ws.max_column - trim_col + 1)
        trimmed = True

    return (shifted, trimmed, cleared)


def run_part_b(wb) -> None:
    for name in wb.sheetnames:
        if not is_location_sheet(name):
            continue
        ws = wb[name]
        _process_sheet_cleanup(ws, name)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    if not INPUT_FILE.exists():
        raise SystemExit(f"Missing input: {INPUT_FILE}")

    wb = openpyxl.load_workbook(str(INPUT_FILE))

    run_part_a(wb)
    run_part_b(wb)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUTPUT_FILE))
    print(f"\n✅ Step 11 complete: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
