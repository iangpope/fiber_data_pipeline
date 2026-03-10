"""Step 4: Color rows within the OPTICAL SPLITTERS section and highlight
first-row-of-each-device-group in yellow.

Reads:  output/Colorized_Cut_Sheet.xlsx
        output/Colored_Connections_Table.xlsx
Writes: output/Colorized_Cut_Sheet_Final_v7_highlighted.xlsx
"""

import os
import openpyxl

from config import OUTPUT_DIR, COLOR, get_fill
from naming_utils import is_location_sheet, safe_fill_hex

CUT_SHEET_PATH   = os.path.join(OUTPUT_DIR, "Colorized_Cut_Sheet.xlsx")
CONNECTIONS_PATH = os.path.join(OUTPUT_DIR, "Colored_Connections_Table.xlsx")
OUTPUT_PATH      = os.path.join(OUTPUT_DIR, "Colorized_Cut_Sheet_Final_v7_highlighted.xlsx")

YELLOW_FILL = get_fill("FFFF00")


# ---------------------------------------------------------------------------
# Load connection → color mapping
# ---------------------------------------------------------------------------

def load_connection_colors():
    wb = openpyxl.load_workbook(CONNECTIONS_PATH)
    ws = wb.active
    color_map = {}
    for row in ws.iter_rows(min_row=2):
        location = row[0].value
        if not location:
            continue
        location = str(location)
        color_map.setdefault(location, {})
        for cell in row[3:]:
            if not cell.value:
                continue
            rgb = safe_fill_hex(cell)
            if rgb:
                color_map[location][cell.value.strip()] = rgb
    return color_map


# ---------------------------------------------------------------------------
# Main coloring pass — within the OPTICAL SPLITTERS section
# ---------------------------------------------------------------------------

def apply_splitter_coloring():
    conn_colors = load_connection_colors()
    wb = openpyxl.load_workbook(CUT_SHEET_PATH)

    for sheetname in wb.sheetnames:
        if not is_location_sheet(sheetname):
            continue

        ws = wb[sheetname]
        colors_for_sheet = conn_colors.get(sheetname, {})
        splitter_mode = False
        header_row = None

        for r in range(7, ws.max_row + 1):
            val_a = str(ws.cell(row=r, column=1).value or "")
            val_b = str(ws.cell(row=r, column=2).value or "")
            val_c = str(ws.cell(row=r, column=3).value or "")
            val_g = str(ws.cell(row=r, column=7).value or "")
            val_h = str(ws.cell(row=r, column=8).value or "")
            val_e = str(ws.cell(row=r, column=5).value or "").strip().upper()
            val_n = str(ws.cell(row=r, column=14).value or "").strip()

            if "OPTICAL SPLITTERS" in val_a.upper():
                splitter_mode = True
                continue

            if splitter_mode and val_e == "CONNECTION":
                header_row = r
                continue

            if splitter_mode and header_row and r > header_row:
                # ---- Color columns A–D by port/device type in the splitter ----
                if "COMMON" in val_c.upper():
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["COMMON"])
                elif "1X32" in val_b.upper():
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["1X32"])
                elif "1X2" in val_b.upper():
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["1X2"])
                elif "MUX" in val_b.upper() and "DEMUX" not in val_b.upper():
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["MUX"])
                elif "DEMUX" in val_b.upper():
                    for c in range(1, 5):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["DEMUX"])

                # ---- Color columns F–O by connected-to device type ----
                if "COMMON" in val_g.upper():
                    for c in range(6, 16):
                        ws.cell(row=r, column=c).fill = get_fill(COLOR["COMMON"])
                else:
                    if "1X32" in val_h.upper():
                        for c in range(6, 16):
                            ws.cell(row=r, column=c).fill = get_fill(COLOR["1X32"])
                    elif "1X2" in val_h.upper():
                        for c in range(6, 16):
                            ws.cell(row=r, column=c).fill = get_fill(COLOR["1X2"])
                    elif "MUX" in val_h.upper() and "DEMUX" not in val_h.upper():
                        for c in range(6, 16):
                            ws.cell(row=r, column=c).fill = get_fill(COLOR["MUX"])
                    elif "DEMUX" in val_h.upper():
                        for c in range(6, 16):
                            ws.cell(row=r, column=c).fill = get_fill(COLOR["DEMUX"])

                # Override F–O with connection-table color if available
                if val_n in colors_for_sheet:
                    for c in range(6, 16):
                        ws.cell(row=r, column=c).fill = get_fill(colors_for_sheet[val_n])

                # Fusion splice → yellow on col E
                if val_e == "<- FUSION ->":
                    ws.cell(row=r, column=5).fill = get_fill(COLOR["FUSION"])

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"✅ Saved updated cut sheet to {OUTPUT_PATH}")


# ---------------------------------------------------------------------------
# Highlighting pass — yellow on first row of each device group
# ---------------------------------------------------------------------------

def highlight_splitter_section():
    wb = openpyxl.load_workbook(OUTPUT_PATH)

    for sheet in wb.sheetnames:
        if not is_location_sheet(sheet):
            continue
        ws = wb[sheet]
        optical_splitter_row = None

        for row in range(1, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=1).value
            if isinstance(cell_val, str) and "OPTICAL SPLITTERS" in cell_val.upper():
                optical_splitter_row = row
                break

        if optical_splitter_row is None:
            continue

        data_start_row = optical_splitter_row + 2  # skip the header row
        prev_val = None
        for r in range(data_start_row, ws.max_row + 1):
            col_a_val = ws.cell(row=r, column=1).value
            if r == data_start_row or col_a_val != prev_val:
                ws.cell(row=r, column=1).fill = YELLOW_FILL
                ws.cell(row=r, column=2).fill = YELLOW_FILL
            prev_val = col_a_val

    wb.save(OUTPUT_PATH)
    print(f"✅ Highlighted sheet saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    apply_splitter_coloring()
    highlight_splitter_section()
