"""
3_colorize.py -- Colorize the cut sheet workbook using direction colors from
the Colored Connections Table, then highlight the OPTICAL SPLITTERS sub-table.

This step combines what were previously two separate scripts into a single
workbook load/save cycle:

  Pass 1 -- Main SHEATHS section coloring:
      Each sheath row is colored by the direction color assigned to its cable
      in step 2. PORT rows in tap sheets are highlighted yellow. Splitter port
      rows (COMMON, OUT 1x32, OUT 1x2, MUX, DEMUX) receive distinct colors
      to distinguish port types at a glance.

  Pass 2 -- OPTICAL SPLITTERS sub-table coloring:
      Each splitter row is colored by device type (COMMON, 1X32, 1X2, MUX,
      DEMUX) using columns B, C, G, and H as the source of truth. The cable
      connection color from step 2 takes priority over the device-type color
      on the right-side columns (F-O).

  Pass 3 -- Device group highlighting:
      The first row of each device group in the OPTICAL SPLITTERS section has
      columns A and B highlighted yellow so it is easy to see where one device
      ends and the next begins.

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


# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
COLORED_CONNECTIONS_PATH = os.path.join(OUTPUT_DIR, "Colored_Connections_Table.xlsx")
CUT_SHEET_PATH           = os.path.join(DATA_DIR,   "cut_sheet.xlsx")
OUTPUT_PATH              = os.path.join(OUTPUT_DIR, "Colorized_Cut_Sheet_Final_v7_highlighted.xlsx")

# Matches a fiber-count prefix at the start of a cable name in legacy MST format
# (e.g. "096CT_RC73E_FT_032_TO_..."). Used to detect MST cables by naming
# pattern when they don't appear in the connections table.
MST_CABLE_RX = re.compile(r"\d{3}CT_")

YELLOW_FILL = get_fill("FFFF00")


# ---------------------------------------------------------------------------
# Load cable-to-color mapping from the Colored Connections Table
# ---------------------------------------------------------------------------

def load_cable_colors() -> dict:
    """
    Build a nested dict mapping  location -> {cable_name -> hex_color}  from
    the Colored Connections Table produced by step 2.

    Each cell in the connections table (columns D onward) contains a cable name
    and may have a directional fill color. safe_fill_hex() is used instead of
    reading the fill directly because openpyxl can return indexed or theme-based
    color objects that do not have a usable RGB string; safe_fill_hex normalizes
    these to a plain 6-character hex string or returns None if the cell is unfilled.

    The resulting dict is used by the colorize functions to look up which color
    to apply to each cable row in the cut sheet.
    """
    wb = openpyxl.load_workbook(COLORED_CONNECTIONS_PATH)
    ws = wb.active
    cable_map: dict = {}

    for row in ws.iter_rows(min_row=2):
        loc = row[0].value
        if not loc:
            continue
        for cell in row[3:]:       # connection columns start at index 3 (column D)
            if not cell.value:
                continue
            rgb = safe_fill_hex(cell)
            if rgb:
                # Store as cable_map[location][cable_name] = hex_color
                cable_map.setdefault(loc, {})[cell.value.strip()] = rgb

    return cable_map


# ---------------------------------------------------------------------------
# Pass 1 -- Color the main SHEATHS section
# ---------------------------------------------------------------------------

def _apply_main_section_colors(ws, conn_dict: dict, splitter_row: int | None) -> None:
    """
    Apply directional and special-purpose colors to all rows in the main
    SHEATHS section (rows 7 through the row before OPTICAL SPLITTERS, or
    to the last row if there is no OPTICAL SPLITTERS section).

    Left-side columns (A-I, col 1-9):
      The cable name is read from column B (col 2). If a color exists in
      conn_dict for that cable, the entire row is filled with that color.
      The first occurrence of each cable name also gets columns A-B filled
      yellow to mark the "splice start" of that sheath.

    Right-side columns (K-T, col 11-20):
      Column O (col 15) is searched for the cable name on the far end of
      the splice. PORT rows receive a yellow fill. For splitter sheets,
      port type colors (COMMON, 1x32, 1x2, MUX, DEMUX) are applied based
      on the value in column Q (device port type) and column S (device name).
    """
    is_splitter = splitter_row is not None
    seen_left: set = set()   # tracks cables already colored on the left side

    end_row = (splitter_row - 1) if splitter_row else ws.max_row

    for r in range(7, end_row + 1):
        # -- Left-side (incoming sheath) --
        val_left = ws.cell(row=r, column=2).value
        if isinstance(val_left, str):
            val_left = val_left.strip()
            fill = conn_dict.get(val_left)

            # Fallback: recognize legacy MST cable naming by pattern
            if not fill and MST_CABLE_RX.match(val_left):
                fill = COLOR["MST"]

            if fill and fill in FILLS:
                # Color the full left-side group (cols A through I).
                for c in range(1, 10):
                    ws.cell(row=r, column=c).fill = FILLS[fill]

                # On the first row for this cable, mark cols A-B yellow to
                # indicate the start of this sheath's splice block.
                if val_left not in seen_left:
                    ws.cell(row=r, column=1).fill = FILLS[COLOR["FUSION"]]
                    ws.cell(row=r, column=2).fill = FILLS[COLOR["FUSION"]]
                    seen_left.add(val_left)

        # -- Center column (J, col 10) -- fusion marker --
        if str(ws.cell(row=r, column=10).value).strip() == "<- FUSION ->":
            ws.cell(row=r, column=10).fill = FILLS[COLOR["FUSION"]]

        # -- Right-side (outgoing sheath) --
        val_right = ws.cell(row=r, column=15).value
        if isinstance(val_right, str):
            val_right = val_right.strip()
            fill = conn_dict.get(val_right)
            if not fill and MST_CABLE_RX.match(val_right):
                fill = COLOR["MST"]
            if fill and fill in FILLS:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[fill]

        # -- PORT rows: override right-side with yellow --
        # PORT rows mark the boundary of a fiber port block in tap sheets.
        # They must be yellow so step 7 can detect and reorder them.
        for c in range(11, 21):
            val = ws.cell(row=r, column=c).value
            if isinstance(val, str) and "PORT" in val.upper() and val.upper() != "PORT NAME":
                for cc in range(11, 21):
                    ws.cell(row=r, column=cc).fill = FILLS[COLOR["FUSION"]]
                break   # stop checking once PORT is found on this row

        # -- Splitter port type colors (right side, splitter sheets only) --
        # Column Q (17) contains the port direction: COMMON, OUT, CH
        # Column S (19) contains the device name: 1X2, 1X32, MUX, DEMUX
        if is_splitter:
            q_val = str(ws.cell(row=r, column=17).value or "").upper().strip()
            s_val = str(ws.cell(row=r, column=19).value or "").upper().strip()

            if "COMMON" in q_val:
                for c in range(11, 21):
                    ws.cell(row=r, column=c).fill = FILLS[COLOR["COMMON"]]
            elif "OUT" in q_val:
                # Distinguish 1x2 (dark pink) from 1x32 (light pink) splitter outputs.
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
# Pass 2 -- Color the OPTICAL SPLITTERS sub-table
# ---------------------------------------------------------------------------

def _apply_optical_splitter_colors(ws, conn_dict: dict, splitter_row: int) -> None:
    """
    Apply device-type colors to the OPTICAL SPLITTERS sub-table.

    Left columns (A-D, col 1-4) are colored by the device type reported in
    column C (device port direction: COMMON, 1X32, 1X2, MUX, DEMUX).

    Right columns (F-O, col 6-15) are colored by the connected-to device type
    reported in columns G and H. If the cable name in column N matches an entry
    in conn_dict (the step 2 color table), the connection color overrides the
    device-type color on the right side.
    """
    header_row = None   # will be set when the CONNECTION header row is found

    for r in range(splitter_row + 1, ws.max_row + 1):
        # Read relevant columns for this row.
        val_b = str(ws.cell(row=r, column=2).value or "")
        val_c = str(ws.cell(row=r, column=3).value or "")
        val_e = str(ws.cell(row=r, column=5).value or "").strip().upper()
        val_g = str(ws.cell(row=r, column=7).value or "")
        val_h = str(ws.cell(row=r, column=8).value or "")
        val_n = str(ws.cell(row=r, column=14).value or "").strip()

        # The first row where column E reads "CONNECTION" is the header row;
        # skip it but note its position so we can start processing data below it.
        if val_e == "CONNECTION":
            header_row = r
            continue

        if header_row is None or r <= header_row:
            continue   # haven't found the header row yet

        # -- Left-side colors (cols A-D): device type of the left device --
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

        # -- Right-side colors (cols F-O): device type of the connected device --
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

        # Override right-side with connection color if available in the color table.
        if val_n in conn_dict:
            for c in range(6, 16):
                ws.cell(row=r, column=c).fill = get_fill(conn_dict[val_n])

        # Mark fusion splices in column E yellow.
        if val_e == "<- FUSION ->":
            ws.cell(row=r, column=5).fill = get_fill(COLOR["FUSION"])


# ---------------------------------------------------------------------------
# Pass 3 -- Highlight the first row of each device group
# ---------------------------------------------------------------------------

def _highlight_device_group_starts(ws, splitter_row: int) -> None:
    """
    Yellow-highlight columns A and B on the first row of each device group
    in the OPTICAL SPLITTERS section.

    Device groups are identified by changes in the column A cell value (which
    holds the device UUID or identifier). When the value in column A changes
    relative to the previous row, that row is the start of a new device group.
    Highlighting the group-start row makes it easy to see at a glance where
    one splitter or MUX ends and the next begins.
    """
    data_start = splitter_row + 2   # row after the OPTICAL SPLITTERS header row
    prev_val = None

    for r in range(data_start, ws.max_row + 1):
        col_a_val = ws.cell(row=r, column=1).value
        # Highlight this row if it is the first data row or the device UUID changed.
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

    for sheetname in wb.sheetnames:
        # Skip non-location sheets (Index, Legend, etc.)
        if not is_location_sheet(sheetname):
            continue

        ws         = wb[sheetname]
        conn_dict  = cable_colors.get(sheetname, {})

        # -- Pre-pass: delete rows that are entirely blank or that only have
        #    values in columns G-I (junk padding columns from the raw export).
        #    Scanning from the bottom prevents row index shifting.
        rows_to_delete = []
        for r in range(ws.max_row, 6, -1):
            row       = [ws.cell(row=r, column=c).value for c in range(1, 21)]
            g_to_i    = row[6:9]        # columns G, H, I (index 6-8)
            non_g_to_i = row[:6] + row[9:]
            # Delete if the row is completely empty, or if only G-I have values.
            if not any(row) or (any(g_to_i) and not any(non_g_to_i)):
                rows_to_delete.append(r)
        for r in rows_to_delete:
            ws.delete_rows(r)

        # Detect the OPTICAL SPLITTERS section row (or None if this sheet has none).
        srow = optical_splitters_row(ws)

        # Run the three color passes.
        _apply_main_section_colors(ws, conn_dict, srow)

        if srow is not None:
            _apply_optical_splitter_colors(ws, conn_dict, srow)
            _highlight_device_group_starts(ws, srow)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"Colorized + highlighted cut sheet saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
