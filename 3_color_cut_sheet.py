import os
import re
import openpyxl

from config import DATA_DIR, OUTPUT_DIR, FILLS, get_fill
from naming_utils import (
    is_location_sheet,
    sheet_has_optical_splitters,
    safe_fill_hex,
)

COLORED_CONNECTIONS_PATH = os.path.join(OUTPUT_DIR, "Colored_Connections_Table.xlsx")
CUT_SHEET_PATH = os.path.join(DATA_DIR, "cut_sheet.xlsx")
COLORIZED_CUT_SHEET_PATH = os.path.join(OUTPUT_DIR, "Colorized_Cut_Sheet.xlsx")

# Pre-compiled regex for MST cable names (e.g. 048CT_...)
MST_CABLE_RX = re.compile(r"\d{3}CT_")

# FILLS and COLOR constants are imported from config.py

def load_cable_colors():
    """Load cable → color mapping from the Colored Connections Table.
    Uses safe_fill_hex() to guard against non-RGB (indexed/theme) fills.
    """
    wb = openpyxl.load_workbook(COLORED_CONNECTIONS_PATH)
    ws = wb.active
    cable_map = {}
    for row in ws.iter_rows(min_row=2):
        loc = row[0].value
        if not loc:
            continue
        for cell in row[3:]:
            if not cell.value:
                continue
            rgb = safe_fill_hex(cell)
            if rgb:
                cable_map.setdefault(loc, {})[cell.value.strip()] = rgb
    return cable_map

def apply_colors_to_cut_sheet():
    color_map = load_cable_colors()
    wb = openpyxl.load_workbook(CUT_SHEET_PATH)

    for sheetname in wb.sheetnames:
        if not is_location_sheet(sheetname):
            continue

        ws = wb[sheetname]
        conn_dict = color_map.get(sheetname, {})

        rows_to_delete = []
        for r in range(ws.max_row, 6, -1):
            row = [ws.cell(row=r, column=c).value for c in range(1, 21)]
            g_to_i = row[6:9]
            non_g_to_i = row[:6] + row[9:]
            if not any(row) or (any(g_to_i) and not any(non_g_to_i)):
                rows_to_delete.append(r)
        for r in rows_to_delete:
            ws.delete_rows(r)

        # Detect splitter sheets by content: does col A contain "OPTICAL SPLITTERS"?
        is_splitter_or_node = any(
            isinstance(ws.cell(row=r2, column=1).value, str)
            and "OPTICAL SPLITTERS" in ws.cell(row=r2, column=1).value.upper()
            for r2 in range(1, ws.max_row + 1)
        )

        seen_left = set()
        for r in range(7, ws.max_row + 1):
            val_left = ws.cell(row=r, column=2).value
            if isinstance(val_left, str):
                val_left = val_left.strip()
                fill = conn_dict.get(val_left)
                if not fill and MST_CABLE_RX.match(val_left):
                    fill = "FF0000"
                if fill and fill in FILLS:
                    for c in range(1, 10):
                        ws.cell(row=r, column=c).fill = FILLS[fill]
                    if val_left not in seen_left:
                        ws.cell(row=r, column=1).fill = FILLS["FFFF00"]
                        ws.cell(row=r, column=2).fill = FILLS["FFFF00"]
                        seen_left.add(val_left)

            val_right = ws.cell(row=r, column=15).value
            if isinstance(val_right, str):
                val_right = val_right.strip()
                fill = conn_dict.get(val_right)
                if not fill and MST_CABLE_RX.match(val_right):
                    fill = "FF0000"
                if fill and fill in FILLS:
                    for c in range(11, 21):
                        ws.cell(row=r, column=c).fill = FILLS[fill]

            if str(ws.cell(row=r, column=10).value).strip() == "<- FUSION ->":
                ws.cell(row=r, column=10).fill = FILLS["FFFF00"]

            for c in range(11, 21):
                val = ws.cell(row=r, column=c).value
                if isinstance(val, str) and "PORT" in val.upper() and val.upper() != "PORT NAME":
                    for cc in range(11, 21):
                        ws.cell(row=r, column=cc).fill = FILLS["FFFF00"]
                    break  # only color once per row if multiple PORT values exist

            q_val = str(ws.cell(row=r, column=17).value).upper().strip() if ws.cell(row=r, column=17).value else ""
            s_val = str(ws.cell(row=r, column=19).value).upper().strip() if ws.cell(row=r, column=19).value else ""

            if is_splitter_or_node:
                if "COMMON" in q_val:
                    for c in range(11, 21):
                        ws.cell(row=r, column=c).fill = FILLS["ADD8E6"]
                elif "OUT" in q_val:
                    if "1X2" in s_val:
                        for c in range(11, 21):
                            ws.cell(row=r, column=c).fill = FILLS["DB7093"]
                    else:
                        for c in range(11, 21):
                            ws.cell(row=r, column=c).fill = FILLS["FFB6C1"]
                elif "CH" in q_val and "MUX" in s_val and "DEMUX" not in s_val:
                    for c in range(11, 21):
                        ws.cell(row=r, column=c).fill = FILLS["FFDAB9"]
                elif "CH" in q_val and "DEMUX" in s_val:
                    for c in range(11, 21):
                        ws.cell(row=r, column=c).fill = FILLS["FFA07A"]

    os.makedirs(os.path.dirname(COLORIZED_CUT_SHEET_PATH), exist_ok=True)
    wb.save(COLORIZED_CUT_SHEET_PATH)
    print(f"✅ Colorized cut sheet saved to {COLORIZED_CUT_SHEET_PATH}")

if __name__ == "__main__":
    apply_colors_to_cut_sheet()
