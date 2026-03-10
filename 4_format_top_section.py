"""
4_format_top_section.py -- Insert a metadata header and direction bar legend
into each location sheet of the colorized cut sheet workbook.

For each location sheet that exists in the Colored Connections Table, this
script:

  1. Removes any previously inserted top section (rows above SHEATHS) so it
     is safe to re-run without duplicating the header.

  2. Inserts a metadata block (rows 2-7):
       A2:A7  -- label column: Splice ID, Enclosure, No. of Trays, Location,
                  Latitude, Longitude
       B2:B7  -- value column filled with pale yellow

  3. Builds a set of colored direction bars (rows 9 onward, one per cable
     direction/type). Each bar spans columns A-E with the direction label,
     cable identifier, and fiber count in specific cells.

  Bar types and their order:
       OLT   -- olive green  (connection to the head-end OLT rack)
       MST   -- red          (main sheath terminal tap)
       North -- orange
       East  -- green
       West  -- slate
       South -- brown
       Splitter -- pink/dark-pink (1x32 / 1x2 splitter devices)
       MUX/DEMUX -- peach/salmon

A debug CSV is written to output/step5_bar_debug.csv for validating color
classification during development. Set DEBUG_ENABLED = False to suppress it.

Reads:  output/Colored_Connections_Table.xlsx
        output/Colorized_Cut_Sheet_Final_v7_highlighted.xlsx
Writes: output/Combined_Formatted_Output.xlsx
        output/step5_bar_debug.csv  (if DEBUG_ENABLED)
"""

import re
import csv
import os
from datetime import datetime
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment

from config import COLOR
from naming_utils import safe_fill_hex


# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
CONN_TABLE_PATH = "output/Colored_Connections_Table.xlsx"
INPUT_WB_PATH   = "output/Colorized_Cut_Sheet_Final_v7_highlighted.xlsx"
OUTPUT_WB_PATH  = "output/Combined_Formatted_Output.xlsx"

DEBUG_DUMP_PATH = "output/step5_bar_debug.csv"
DEBUG_ENABLED   = True   # write classification details to CSV for review


# ---------------------------------------------------------------------------
# Constants and color mappings
# ---------------------------------------------------------------------------

# Regex to recognize a bare OLT site token (e.g. RC73E, MS90E).
OLT_TOKEN_RX = re.compile(r"^[A-Z]{2}\d{2,3}E$", flags=re.IGNORECASE)

# Hex values for special bar types (pulled from the central config).
COLOR_SPLIT_1X2  = COLOR["1X2"]    # dark pink  -- 1x2 splitter bar
COLOR_SPLIT_1X32 = COLOR["1X32"]   # light pink -- 1x32 splitter bar
COLOR_MUX        = COLOR["MUX"]    # peach      -- MUX bar
COLOR_DEMUX      = COLOR["DEMUX"]  # salmon     -- DEMUX bar
LABEL_FILL       = "FFF9DB"        # pale yellow -- metadata value cell background
OLT_BAR_COLOR    = COLOR["OLT"]    # olive green -- OLT connection bar

# Map from fill hex back to a direction label used when classifying cable colors.
# Only the four cardinal directions, OLT, and MST are mapped.
COLOR_TO_DIRECTION = {
    COLOR["NORTH"]: "North",
    COLOR["SOUTH"]: "South",
    COLOR["EAST"]:  "East",
    COLOR["WEST"]:  "West",
    OLT_BAR_COLOR:  "OLT",
    COLOR["MST"]:   "MST",
}

# Display order for the direction bars in each sheet's top section.
# Lower numbers appear first. Directions not listed default to order 99.
BAR_ORDER = {
    "OLT":      0,
    "MST":      1,
    "North":    2,
    "East":     3,
    "West":     4,
    "South":    5,
    "Splitter": 6,
    "MUX":      7,
    "DEMUX":    8,
}

# Reference palette for nearest-color classification of cable fill colors.
# Stored as 6-char hex RGB for direct comparison.
PALETTE_RGB = {
    "North": COLOR["NORTH"],
    "South": COLOR["SOUTH"],
    "East":  COLOR["EAST"],
    "West":  COLOR["WEST"],
    "MST":   COLOR["MST"],
}


# ---------------------------------------------------------------------------
# Color utility functions
# ---------------------------------------------------------------------------

def _hex_to_rgb(h: str):
    """
    Convert a 6-character hex string to an (R, G, B) integer tuple.
    Returns None if the string is not a valid 6-character hex color.
    """
    h = (h or "").strip().upper()
    h = h[-6:] if len(h) >= 6 else h
    if len(h) != 6 or not all(ch in "0123456789ABCDEF" for ch in h):
        return None
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def classify_direction_from_color(rgb6: str):
    """
    Map a 6-character hex color to the nearest direction label in PALETTE_RGB.

    Uses squared Euclidean distance in RGB space. A conservative threshold
    (120 per channel) is applied so that unrecognized theme or accent colors
    are not mislabeled -- they return None instead.

    Returns 'North', 'East', 'South', 'West', 'MST', or None.
    """
    rgb = _hex_to_rgb(rgb6)
    if rgb is None:
        return None
    best = None
    for name, pal_hex in PALETTE_RGB.items():
        prgb = _hex_to_rgb(pal_hex)
        if prgb is None:
            continue
        # Squared Euclidean distance in RGB space.
        d = (rgb[0] - prgb[0])**2 + (rgb[1] - prgb[1])**2 + (rgb[2] - prgb[2])**2
        if best is None or d < best[0]:
            best = (d, name)
    # Only accept the match if it is within the tolerance threshold.
    if best and best[0] <= 120**2:
        return best[1]
    return None


def palette_distance_report(rgb6: str):
    """
    Return (best_name, best_dist_sq, dist_sq_map) for debugging.

    Computes the squared RGB distance from rgb6 to every palette color and
    returns the closest match along with the full distance map. Used only
    when DEBUG_ENABLED is True to populate the debug CSV.
    """
    rgb = _hex_to_rgb(rgb6)
    if rgb is None:
        return (None, None, {})
    dist_map  = {}
    best_name = None
    best_d    = None
    for name, pal_hex in PALETTE_RGB.items():
        prgb = _hex_to_rgb(pal_hex)
        if prgb is None:
            continue
        d = (rgb[0] - prgb[0])**2 + (rgb[1] - prgb[1])**2 + (rgb[2] - prgb[2])**2
        dist_map[name] = d
        if best_d is None or d < best_d:
            best_d    = d
            best_name = name
    return (best_name, best_d, dist_map)


def to_argb(rgb6: str) -> str:
    """
    Convert a 6-character RGB hex string to an 8-character ARGB string with
    full opacity (FF alpha prefix). openpyxl's PatternFill requires ARGB format.
    """
    rgb6 = (rgb6 or "").strip().upper()[-6:]
    return "FF" + rgb6


# ---------------------------------------------------------------------------
# Worksheet helper functions
# ---------------------------------------------------------------------------

def parse_connection_value(val: str):
    """
    Parse a cable connection value string into (fiber_count, endpoint_A, endpoint_B).

    Supports both naming conventions:
      Old: '48CT RC73E_FT_001 TO RC73E_SE_001'
      New: 'RC73E_FT_001_TO_RC73E_SE_001_48CT'

    Returns (fiber, locA, locB) on success, or (None, None, None) if unparseable.
    """
    t = str(val or "").strip()
    if not t:
        return None, None, None

    # Old format: fiber count comes first, separated by spaces.
    m = re.match(r"^(\d{2,3}CT)\s+(\S+)\s+TO\s+(\S+)$", t, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper(), m.group(2).strip(), m.group(3).strip()

    # New format: "_TO_" delimiter; fiber count is a suffix on the second endpoint.
    if "_TO_" in t:
        a, b = t.split("_TO_", 1)
        fiber = None
        m2 = re.search(r"_(\d{2,3}CT)\b", b, flags=re.IGNORECASE)
        if m2:
            fiber = m2.group(1).upper()
            b = re.sub(r"_(\d{2,3}CT)\b", "", b, flags=re.IGNORECASE)
        return fiber, a.strip(), b.strip()

    return None, None, None


def cell_fill_hex(cell) -> str | None:
    """
    Return the last-6-character RGB hex for a cell with a solid fill.
    Returns None for cells with no fill or non-solid fill patterns.

    This is a local duplicate of safe_fill_hex for cells read from the
    connections table workbook where the fill was set programmatically.
    """
    f = getattr(cell, "fill", None)
    if not f or getattr(f, "patternType", None) != "solid":
        return None
    sc = getattr(f, "start_color", None)
    if sc is not None:
        rgb = getattr(sc, "rgb", None) or getattr(sc, "index", None)
    if isinstance(rgb, str):
        rgb = rgb.strip().upper()
        return rgb[-6:]   # strip FF alpha prefix if present
    return None


def find_row_in_colA(ws, text: str) -> int | None:
    """
    Return the row number of the first cell in column A whose string value
    exactly matches text, or None if not found.
    """
    for cell in ws["A"]:
        if str(cell.value).strip() == text:
            return cell.row
    return None


def guess_olt_token(sheet_name: str) -> str:
    """
    Derive the OLT site token from a sheet name.

    For bare OLT sheets (e.g. 'RC73E'), the name itself is the token.
    For location sheets (e.g. 'RC73E_FT_001'), the prefix before the first
    underscore is the token. Returns empty string if no token can be found.
    """
    s = str(sheet_name).strip()
    if OLT_TOKEN_RX.match(s):
        return s
    prefix = s.split("_", 1)[0].strip()
    if OLT_TOKEN_RX.match(prefix):
        return prefix
    return ""


# ---------------------------------------------------------------------------
# Load connection table data
# ---------------------------------------------------------------------------

conn_wb    = load_workbook(CONN_TABLE_PATH)
conn_sheet = conn_wb["Colored Connections"]

# conn_dict: location -> list of (fiber, locA, locB, fill_hex) for each cable
# coords:    location -> (latitude, longitude)
conn_dict: dict = {}
coords:    dict = {}

for row in conn_sheet.iter_rows(min_row=2):   # row 1 is the header
    loc = row[0].value
    if loc is None:
        continue

    lat = row[1].value
    lon = row[2].value
    coords[str(loc).strip()] = (lat, lon)

    connections = []
    for cell in row[3:]:   # connection columns start at index 3 (column D)
        val = cell.value
        if not val:
            continue
        fiber, locA, locB = parse_connection_value(val)
        if not fiber:
            continue
        # safe_fill_hex guards against indexed/theme colors from the connections table
        fill_color = safe_fill_hex(cell)
        connections.append((fiber, str(locA), str(locB), (fill_color or "").upper()))

    conn_dict[str(loc).strip()] = connections


# ---------------------------------------------------------------------------
# Load the colorized cut sheet workbook
# ---------------------------------------------------------------------------

wb = load_workbook(INPUT_WB_PATH)

# Shared style objects used in the metadata block.
yellow_fill = PatternFill(start_color=LABEL_FILL, end_color=LABEL_FILL, fill_type="solid")
white_fill  = PatternFill(start_color="FFFFFF",   end_color="FFFFFF",   fill_type="solid")

thin_side  = Side(border_style="thin",  color="000000")
thick_side = Side(border_style="thick", color="000000")

# Accumulate debug classification rows for the CSV export.
debug_rows = []
run_id     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Process each sheet
# ---------------------------------------------------------------------------

for sheet_name in wb.sheetnames:
    loc_key = str(sheet_name).strip()
    if loc_key not in conn_dict:
        continue   # sheet not in connections table; skip

    ws = wb[sheet_name]

    # Remove any previously inserted top section so re-running is safe.
    sheaths_row = find_row_in_colA(ws, "SHEATHS")
    if sheaths_row and sheaths_row > 1:
        ws.delete_rows(1, sheaths_row - 1)

    # Determine the OLT site token for this sheet (used to detect OLT cables).
    olt_id = guess_olt_token(loc_key)

    # Build the GPS coordinate strings for display in the metadata block.
    lat_val, lon_val = coords.get(loc_key, (None, None))
    lat_str = f"{lat_val:.5f}" if isinstance(lat_val, (float, int)) else (lat_val or "")
    lon_str = f"{lon_val:.5f}" if isinstance(lon_val, (float, int)) else (lon_val or "")

    # Metadata label/value pairs for the A2:B7 block.
    labels = ["Splice ID:", "Enclosure:", "No. of Trays:", "Location:", "Latitude:", "Longitude:"]
    values = [loc_key, "", "", "", lat_str, lon_str]

    # ------------------------------------------------------------------
    # Build direction bars from the connection entries for this location.
    #
    # Bars are grouped by (direction_label, fiber_count, fill_color) so that
    # cables going the same direction with the same fiber count are combined
    # into one bar rather than creating a bar per cable.
    # ------------------------------------------------------------------
    current_conns = conn_dict[loc_key]
    grouped = {}   # (label, fiber, fill_rgb6) -> count

    for fiber, locA, locB, color in current_conns:
        color = (color or "").upper()
        if not color:
            continue

        # OLT override: if either endpoint is the bare OLT site token, this
        # cable connects directly to the OLT rack and gets the olive-green bar.
        if olt_id and (
            str(locA).upper() == str(olt_id).upper() or
            str(locB).upper() == str(olt_id).upper()
        ):
            label     = "OLT"
            fill_rgb6 = OLT_BAR_COLOR
        else:
            # Classify the directional fill color to a compass label.
            label = classify_direction_from_color(color)
            if not label:
                continue   # unrecognized color; skip this cable
            fill_rgb6 = color[-6:]

        # Optional: record detailed classification data for the debug CSV.
        if DEBUG_ENABLED:
            best_name, best_d, dist_map = palette_distance_report(color)
            debug_rows.append({
                "run_id":           run_id,
                "sheet":            loc_key,
                "olt_id":           olt_id,
                "fiber":            fiber,
                "locA":             str(locA),
                "locB":             str(locB),
                "raw_fill":         color.upper(),
                "olt_override":     "1" if (
                    olt_id and (
                        str(locA).upper() == str(olt_id).upper() or
                        str(locB).upper() == str(olt_id).upper()
                    )
                ) else "0",
                "classified_label": label,
                "bar_fill":         fill_rgb6,
                "best_palette":     best_name,
                "best_dist_sq":     best_d,
                "dist_North":       dist_map.get("North"),
                "dist_South":       dist_map.get("South"),
                "dist_East":        dist_map.get("East"),
                "dist_West":        dist_map.get("West"),
                "dist_MST":         dist_map.get("MST"),
            })

        key           = (label, fiber, fill_rgb6)
        grouped[key]  = grouped.get(key, 0) + 1

    # Format grouped entries into bar tuples: (label, count, textB, fiber, fill_rgb6)
    connection_bars = []
    for (label, fiber, fill_rgb6), n in grouped.items():
        # textB is the OLT site token for OLT bars; None for directional bars.
        textB = olt_id if label == "OLT" else None
        connection_bars.append((label, n, textB, fiber, fill_rgb6))

    # ------------------------------------------------------------------
    # Build splitter and MUX/DEMUX bars by scanning the OPTICAL SPLITTERS
    # section of the current sheet. One bar is added per unique device type.
    # ------------------------------------------------------------------
    splitter_bars = []
    mux_bars      = []

    found_1x32 = found_1x2 = Found_mux = found_demux = False

    opt_row = find_row_in_colA(ws, "OPTICAL SPLITTERS")
    if opt_row:
        r = opt_row + 1
        while True:
            valA = ws.cell(row=r, column=1).value
            valB = ws.cell(row=r, column=2).value
            if not valA and not valB:
                break   # end of the OPTICAL SPLITTERS sub-table

            if isinstance(valB, str):
                text = valB.upper()
                # Add one bar per unique device type found in the sub-table.
                if (not found_1x32) and ("_1X32" in text):
                    splitter_bars.append(("Splitter", 1, None, "1X32", COLOR_SPLIT_1X32))
                    found_1x32 = True
                if (not found_1x2) and ("_1X2" in text):
                    splitter_bars.append(("Splitter", 1, None, "1X2", COLOR_SPLIT_1X2))
                    found_1x2 = True
                if (not found_demux) and ("DEMUX" in text):
                    subtype = "40CH" if "40CH" in text else "4CH"
                    mux_bars.append(("DEMUX", 1, None, subtype, COLOR_DEMUX))
                    found_demux = True
                if (not Found_mux) and ("MUX" in text):
                    subtype = "40CH" if "40CH" in text else "4CH"
                    mux_bars.append(("MUX", 1, None, subtype, COLOR_MUX))
                    Found_mux = True
            r += 1

    # Sort connection bars by the defined display order, then by fiber/label.
    connection_bars.sort(
        key=lambda b: (BAR_ORDER.get(b[0], 99), str(b[3] or ""), str(b[2] or ""))
    )

    # Final ordered bars list: direction bars, then splitter bars, then MUX/DEMUX.
    bars = connection_bars + splitter_bars + mux_bars

    # Record bar output in debug CSV.
    if DEBUG_ENABLED:
        for i, (lbl, cnt, textB, fiber_or_kind, fill_hex) in enumerate(bars):
            debug_rows.append({
                "run_id":       run_id,
                "sheet":        loc_key,
                "record_type":  "bar",
                "bar_index":    i,
                "bar_label":    lbl,
                "bar_count":    cnt,
                "bar_textB":    textB or "",
                "bar_fiber":    fiber_or_kind or "",
                "bar_fill":     fill_hex,
            })

    # ------------------------------------------------------------------
    # Insert the top section rows into the sheet.
    #
    # Layout:
    #   Row 1     : blank
    #   Rows 2-7  : metadata block (A:B)
    #   Row 8     : blank
    #   Rows 9+   : one bar per direction/type
    #   Final row : blank
    # ------------------------------------------------------------------
    total_new_rows = 1 + 6 + 1 + len(bars) + 1
    ws.insert_rows(1, total_new_rows)

    # Write the metadata label/value pairs into A2:B7.
    meta_start = 2
    for i, (lab, val) in enumerate(zip(labels, values)):
        r  = meta_start + i
        cA = ws.cell(row=r, column=1)
        cB = ws.cell(row=r, column=2)
        cA.value = lab
        cB.value = val

        # Label column: bold text, white background, left-aligned.
        cA.font      = Font(bold=True)
        cA.alignment = Alignment(horizontal="left")
        cA.fill      = white_fill

        # Value column: centered, pale-yellow background.
        cB.alignment = Alignment(horizontal="center")
        cB.fill      = yellow_fill

    meta_end = meta_start + len(labels) - 1   # last metadata row (row 7)

    # Apply a thick outer border with thin interior seam to the metadata block.
    for r in range(meta_start, meta_end + 1):
        for c in (1, 2):
            left   = thick_side if c == 1 else thin_side
            right  = thick_side if c == 2 else thin_side
            top    = thick_side if r == meta_start else thin_side
            bottom = thick_side if r == meta_end   else thin_side
            # The seam between column A and B should always be thin.
            if c == 1:
                right = thin_side
            else:
                left = thin_side
            ws.cell(row=r, column=c).border = Border(
                left=left, right=right, top=top, bottom=bottom
            )

    start_bar_row = meta_end + 2   # bars begin at row 9

    def set_bar_cell(cell, fill_hex, value=None) -> None:
        """
        Apply the bar style to a single cell: solid fill, bold white text,
        centered alignment.
        """
        cell.value     = value
        cell.fill      = PatternFill(
            start_color=to_argb(fill_hex),
            end_color=to_argb(fill_hex),
            fill_type="solid",
        )
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Write each bar across columns A-E.
    for offset, (label, count, textB, fiber_or_kind, fill_hex) in enumerate(bars):
        # Column A: label with count when more than one cable (e.g. "North (2)").
        # OLT bars show just "OLT" without a count.
        textA = (
            f"{label} ({count})"
            if count > 1 and label not in ("OLT",)
            else label
        )
        textD = fiber_or_kind   # fiber count or device subtype in column D
        r     = start_bar_row + offset

        # Fill columns A through E with the bar color and labels.
        for c in range(1, 6):
            val = None
            if c == 1:
                val = textA          # direction/type label
            elif c == 2 and textB:
                val = textB          # OLT site token (OLT bars only)
            elif c == 4 and textD:
                val = textD          # fiber count or device subtype
            set_bar_cell(ws.cell(row=r, column=c), fill_hex, val)


# ---------------------------------------------------------------------------
# Save output workbook
# ---------------------------------------------------------------------------

wb.save(OUTPUT_WB_PATH)

# ---------------------------------------------------------------------------
# Write debug CSV (classification audit trail)
# ---------------------------------------------------------------------------

if DEBUG_ENABLED:
    try:
        os.makedirs(os.path.dirname(DEBUG_DUMP_PATH), exist_ok=True)
        fieldnames = [
            "run_id", "record_type", "sheet", "olt_id", "fiber", "locA", "locB",
            "raw_fill", "olt_override", "classified_label", "bar_fill",
            "best_palette", "best_dist_sq",
            "dist_North", "dist_South", "dist_East", "dist_West", "dist_MST",
            "bar_index", "bar_label", "bar_count", "bar_textB", "bar_fiber",
        ]
        with open(DEBUG_DUMP_PATH, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            for row in debug_rows:
                # Tag connection-classification rows (vs. bar output rows).
                if "record_type" not in row:
                    row["record_type"] = "conn"
                w.writerow(row)
        print(f"Debug dump: {DEBUG_DUMP_PATH}")
    except Exception as e:
        print(f"Debug dump failed: {e}")

print(f"Wrote: {OUTPUT_WB_PATH}")