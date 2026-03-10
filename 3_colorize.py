"""Step 3: Colorize the cut sheet and highlight the OPTICAL SPLITTERS section.

Replaces the two-step process of running 3_color_cut_sheet.py followed by
4_update_and_highlight_splitter_section.py.

Reads:  data/cut_sheet.xlsx
        output/Colored_Connections_Table.xlsx
Writes: output/Colorized_Cut_Sheet_Final_v7_highlighted.xlsx
"""

import os
import re
import openpyxl

from config import DATA_DIR, OUTPUT_DIR, COLOR, FILLS, get_fill
from naming_utils import (
    is_location_sheet,
    sheet_has_optical_splitters,
    optical_splitters_row,
    safe_fill_hex,
)

COLORED_CONNECTIONS_PATH = os.path.join(OUTPUT_DIR, "Colored_Connections_Table.xlsx")
CUT_SHEET_PATH           = os.path.join(DATA_DIR,   "cut_sheet.xlsx")
OUTPUT_PATH              = os.path.join(OUTPUT_DIR, "Colorized_Cut_Sheet_Final_v7_highlighted.xlsx")

MST_CABLE_RX = re.compile(r"\d{3}CT_")
YELLOW_FILL  = get_fill("FFFF00")


# ---------------------------------------------------------------------------
# Load connection → color mapping
# ---------------------------------------------------------------------------

def load_cable_colors() -> dict:
    """Load cable → color mapping from the Colored Connections Table.
    Uses safe_fill_hex() to guard against non-RGB (indexed/theme) fills.
    """
    wb = openpyxl.load_workbook(COLORED_CONNECTIONS_PATH)
    ws = wb.active
    cable_map: dict = {}
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


# ---------------------------------------------------------------------------
# Pass 1: Color the main sheath section (left + right sides)
# ---------------------------------------------------------------------------

def _apply_main_section_colors(ws, conn_dict: dict, splitter_row: int | None) -> None:
    """Color sheath rows and splitter output port rows in the SHEATHS section."""
    is_splitter = splitter_row is not None
    seen_left: set = set()

    end_row = (splitter_row - 1) if splitter_row else ws.max_row

    for r in range(7, end_row + 1):
        val_left = ws.cell(row=r, column=2).value
        if isinstance(val_left, str):
            val_left = val_left.strip()
            fill = conn_dict.get(val_left)
            if not fill and MST_CABLE_RX.match(val_left):
                fill = COLOR["MST"]
            if fill and fill in FILLS:
                for c in range(1, 10):
                    ws.cell(row=r, column=c).fill = FILLS[fill]
                if val_left not in seen_left:
                    ws.cell(row=r, column=1).fill = FILLS[COLOR["FUSION"]]
                    ws.cell(row=r, column=2).fill = FILLS[COLOR["FUSION"]]
                    seen_left.add(val_left)

        val_right = ws.cell(row=r, column=15).value
        if isinstance(val_right, str):
            val_right = val_right.strip()
            fill = conn_dict.get(val_right)
            if not fill and MST_CABLE_RX.match(val_right):
                fill = COLOR["MST"]
            if fill and fill in FILLS:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[fill]

        if str(ws.cell(row=r, column=10).value).strip() == "<- FUSION ->":
            ws.cell(row=r, column=10).fill = FILLS[COLOR["FUSION"]]

        # PORT rows → yellow (only color once per row)
        for c in range(11, 21):
            val = ws.cell(row=r, column=c).value
            if isinstance(val, str) and "PORT" in val.upper() and val.upper() != "PORT NAME":
                for cc in range(11, 21):
                    ws.cell(row=r, column=cc).fill = FILLS[COLOR["FUSION"]]
                break

        # Splitter output port coloring (right-side cols, rows above OPTICAL SPLITTERS)
        if is_splitter:
            q_val = str(ws.cell(row=r, column=17).value or "").upper().strip()
            s_val = str(ws.cell(row=r, column=19).value or "").upper().strip()
            if "COMMON" in q_val:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[COLOR["COMMON"]]
            elif "OUT" in q_val:
                key = "1X2" if "1X2" in s_val else "1X32"
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[COLOR[key]]
            elif "CH" in q_val and "MUX" in s_val and "DEMUX" not in s_val:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[COLOR["MUX"]]
            elif "CH" in q_val and "DEMUX" in s_val:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[COLOR["DEMUX"]]


# ---------------------------------------------------------------------------
# Pass 2: Color the OPTICAL SPLITTERS section
# ---------------------------------------------------------------------------

def _apply_optical_splitter_colors(ws, conn_dict: dict, splitter_row: int) -> None:
    """Color rows within the OPTICAL SPLITTERS sub-table."""
    header_row = None

    for r in range(splitter_row + 1, ws.max_row + 1):
        val_a = str(ws.cell(row=r, column=1).value or "")
        val_b = str(ws.cell(row=r, column=2).value or "")
        val_c = str(ws.cell(row=r, column=3).value or "")
        val_e = str(ws.cell(row=r, column=5).value or "").strip().upper()
        val_g = str(ws.cell(row=r, column=7).value or "")
        val_h = str(ws.cell(row=r, column=8).value or "")
        val_n = str(ws.cell(row=r, column=14).value or "").strip()

        if val_e == "CONNECTION":
            header_row = r
            continue

        if header_row is None or r <= header_row:
            continue

        # Color A–D by device type
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

        # Color F–O by connected-to device type
        if "COMMON" in val_g.upper():
            for c in range(6, 16):
                ws.cell(row=r, column=c).fill = get_fill(COLOR["COMMON"])
        elif "1X32" in val_h.upper():
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
        if val_n in conn_dict:
            for c in range(6, 16):
                ws.cell(row=r, column=c).fill = get_fill(conn_dict[val_n])

        # Fusion splice → yellow on col E
        if val_e == "<- FUSION ->":
            ws.cell(row=r, column=5).fill = get_fill(COLOR["FUSION"])


# ---------------------------------------------------------------------------
# Pass 3: Highlight first row of each device group in yellow
# ---------------------------------------------------------------------------

def _highlight_device_group_starts(ws, splitter_row: int) -> None:
    """Yellow-highlight cols A & B whenever the device UUID (col A) changes."""
    data_start = splitter_row + 2  # skip the header row
    prev_val = None
    for r in range(data_start, ws.max_row + 1):
        col_a_val = ws.cell(row=r, column=1).value
        if r == data_start or col_a_val != prev_val:
            ws.cell(row=r, column=1).fill = YELLOW_FILL
            ws.cell(row=r, column=2).fill = YELLOW_FILL
        prev_val = col_a_val


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> None:
    cable_colors = load_cable_colors()
    wb = openpyxl.load_workbook(CUT_SHEET_PATH)

    # Delete rows that are entirely blank or only have cols G–I filled
    for sheetname in wb.sheetnames:
        if not is_location_sheet(sheetname):
            continue
        ws = wb[sheetname]
        conn_dict = cable_colors.get(sheetname, {})

        rows_to_delete = []
        for r in range(ws.max_row, 6, -1):
            row = [ws.cell(row=r, column=c).value for c in range(1, 21)]
            g_to_i = row[6:9]
            non_g_to_i = row[:6] + row[9:]
            if not any(row) or (any(g_to_i) and not any(non_g_to_i)):
                rows_to_delete.append(r)
        for r in rows_to_delete:
            ws.delete_rows(r)

        # Detect OPTICAL SPLITTERS row (content-based, not name-based)
        srow = optical_splitters_row(ws)

        _apply_main_section_colors(ws, conn_dict, srow)

        if srow is not None:
            _apply_optical_splitter_colors(ws, conn_dict, srow)
            _highlight_device_group_starts(ws, srow)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"✅ Colorized + highlighted cut sheet saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
