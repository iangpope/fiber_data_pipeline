"""
2_compute_all_directions.py -- Compute cable bearing directions and assign colors.

Reads the Connections Table (built by step 1) and the KMZ geometry to determine
which compass direction (North/East/South/West) each fiber cable travels when
leaving a splice enclosure. The result is an Excel workbook where each cable's
cell is filled with a directional color. This file is the manual checkpoint:
review it to verify direction assignments before running steps 3-8.

Special cases:
  - OLT cables (one endpoint is the bare site token, e.g. RC73E) -> olive green
  - MST tap cables (single-connection FT tap nearest to its SE enclosure) -> red

Reads:  data/Connections_Table.xlsx
        data/<project>.kmz
Writes: output/Colored_Connections_Table.xlsx
"""

import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from geopy.distance import geodesic
from math import atan2, degrees, radians, sin, cos
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from naming_utils import parse_location_id


# ---------------------------------------------------------------------------
# Regex patterns for parsing cable names and location tokens.
#
# New-format cable names embed location tokens like "RC73E_FT_032" or
# "RC73E_SE_006". Legacy format uses "MIC..." prefixed tokens. Both styles
# appear in older projects so both patterns are needed.
# ---------------------------------------------------------------------------

# Matches a location token in new naming convention (e.g. RC73E_FT_001)
NEW_LOC_TOKEN_RX    = re.compile(r"[A-Z0-9]+E_(?:FT|SE)_\d{3}")

# Matches a location token in legacy naming convention (e.g. MICRCTS001)
LEGACY_LOC_TOKEN_RX = re.compile(r"\bMIC\w+\b")

# Matches a bare OLT site token (e.g. RC73E, MS90E) -- 2 letters, 2-3 digits, E
OLT_TOKEN_RX        = re.compile(r"^[A-Z]{2}\d{2,3}E$")

# Matches a fiber count suffix in a cable name (e.g. _48CT, _096CT)
FIBER_SUFFIX_RX     = re.compile(r"_(\d{2,3}CT)\b", flags=re.IGNORECASE)


# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
DATA_DIR         = "data"
CONNECTIONS_PATH = os.path.join(DATA_DIR, "Connections_Table.xlsx")
OUTPUT_PATH      = "output/Colored_Connections_Table.xlsx"


def detect_kmz_file() -> str:
    """Return the path to the one KMZ file in the data folder, or raise."""
    for f in os.listdir(DATA_DIR):
        if f.lower().endswith(".kmz"):
            return os.path.join(DATA_DIR, f)
    raise FileNotFoundError("No .kmz file found in 'data' folder.")

KMZ_PATH = detect_kmz_file()


# ---------------------------------------------------------------------------
# Cable name parsing helpers
# ---------------------------------------------------------------------------

def parse_endpoints(cable_name: str):
    """
    Extract (endpointA, endpointB, fiber_count) from a cable name string.

    Supports two naming formats:
      - New format: "RC73E_FT_001_TO_RC73E_SE_001_48CT"  (underscore-delimited)
      - Old format: "48CT RC73E_FT_001 TO RC73E_SE_001"  (space-delimited with fiber first)

    Returns (None, None, None) if the name does not match either pattern.
    """
    t = str(cable_name or "").strip()
    if not t:
        return None, None, None

    # New format: split on "_TO_"
    if "_TO_" in t:
        a, b = t.split("_TO_", 1)
        fiber = None
        m = re.search(r"_(\d{2,3}CT)\b", b, flags=re.IGNORECASE)
        if m:
            fiber = m.group(1).upper()
            b = re.sub(r"_(\d{2,3}CT)\b", "", b, flags=re.IGNORECASE)
        return a.strip(), b.strip(), fiber

    # Old format: "<fiber> <endA> TO <endB>"
    m = re.match(r"^\s*(\d{2,3}CT)\s+(\S+)\s+TO\s+(\S+)\s*$", t, flags=re.IGNORECASE)
    if m:
        return m.group(2).strip(), m.group(3).strip(), m.group(1).upper()

    return None, None, None


def is_olt_cable(cable_name: str, olt_token: str) -> bool:
    """
    Return True if either endpoint of the cable is the bare OLT site token.

    The OLT token is typically the short site identifier (e.g. "RC73E") that
    identifies the head-end optical line terminal rack. Cables connected to
    the OLT receive a distinct color to make the OLT connection obvious.
    """
    a, b, _ = parse_endpoints(cable_name)
    if not olt_token:
        return False
    return (a == olt_token) or (b == olt_token)


# ---------------------------------------------------------------------------
# KMZ parsing: extract all fiber cable line segments
# ---------------------------------------------------------------------------

def extract_fiber_segments() -> list:
    """
    Parse the KMZ and return a list of (cable_name, [(lat, lon), ...]) tuples.

    Each fiber cable in the KMZ is a LineString Placemark. This function
    iterates all placemarks in the document, skipping point placemarks (which
    represent splice enclosure locations, not cables), and collects the name
    and coordinate sequence of every line placemark.

    The KML namespace is stripped so the parser works regardless of whether
    the KML was exported with or without the opengis namespace prefix.
    """
    def strip_ns(tag: str) -> str:
        """Remove the XML namespace prefix from a tag string."""
        return tag.split("}")[-1] if "}" in tag else tag

    with zipfile.ZipFile(KMZ_PATH, "r") as zf:
        kml_file = next((f for f in zf.namelist() if f.endswith(".kml")), None)
        root = ET.fromstring(zf.read(kml_file))

    def recursive_find(el) -> list:
        """Walk the element tree and collect all LineString placemarks."""
        segments = []
        for child in el.iter():
            tag = strip_ns(child.tag.lower())
            if tag != "placemark":
                continue
            coords_elem = child.find(".//{*}LineString/{*}coordinates")
            if coords_elem is None:
                continue   # skip point placemarks (splice enclosure pins)
            # Parse coordinate text: space-separated "lon,lat[,alt]" tokens
            coords  = coords_elem.text.strip().split()
            latlons = [
                (float(c.split(",")[1]), float(c.split(",")[0]))
                for c in coords if "," in c
            ]
            name_elem = next(
                (c for c in child if strip_ns(c.tag) == "name"), None
            )
            name = (
                name_elem.text.strip()
                if (name_elem is not None and name_elem.text)
                else None
            )
            segments.append((name, latlons))
        return segments

    segments = recursive_find(root)
    print(f"Extracted {len(segments)} line segments from KMZ")
    return segments


def merge_named_segments(segments: list) -> list:
    """
    Merge multi-part cable segments that share the same name into single chains.

    GIS exports can split a single cable into multiple LineString segments
    (e.g. when it crosses a tile boundary). This function groups all segments
    by cable name and chains them together end-to-end using a deque. If
    segments cannot be joined geometrically (no shared endpoints), they are
    appended in order to prevent data loss.

    Returns a list of (cable_name, [lat_lon, ...]) with one entry per cable.
    """
    from collections import defaultdict, deque

    # Group all coordinate sequences by cable name.
    grouped = defaultdict(list)
    for name, coords in segments:
        if name:
            grouped[name].append(coords)

    merged = []
    for name, paths in grouped.items():
        if len(paths) == 1:
            # Only one segment for this cable; no merging needed.
            merged.append((name, paths[0]))
            continue

        # Build a chain by connecting segments end-to-end.
        chain = deque(paths[0])
        remaining = list(paths[1:])

        while remaining:
            progress = False
            still_remaining = []
            for p in remaining:
                if chain[-1] == p[0]:           # tail of chain matches head of segment
                    chain.extend(p[1:])
                    progress = True
                elif chain[-1] == p[-1]:         # tail matches tail; append reversed
                    chain.extend(reversed(p[:-1]))
                    progress = True
                elif chain[0] == p[-1]:          # head of chain matches tail; prepend
                    chain.extendleft(reversed(p[:-1]))
                    progress = True
                elif chain[0] == p[0]:           # head matches head; prepend reversed
                    chain.extendleft(p[1:])
                    progress = True
                else:
                    still_remaining.append(p)
            if not progress:
                # No geometric connection found; append remaining segments anyway.
                for p in still_remaining:
                    chain.extend(p)
                break
            remaining = still_remaining

        merged.append((name, list(chain)))

    return merged


# ---------------------------------------------------------------------------
# Bearing calculation and color assignment
# ---------------------------------------------------------------------------

def bearing(p1: tuple, p2: tuple) -> float:
    """
    Compute the geodesic forward bearing from p1 to p2 in degrees (0-360).

    Uses the standard spherical haversine bearing formula. Both points are
    (lat, lon) tuples in decimal degrees. 0 = North, 90 = East, etc.
    """
    lat1, lon1 = radians(p1[0]), radians(p1[1])
    lat2, lon2 = radians(p2[0]), radians(p2[1])
    dlon = lon2 - lon1
    x = sin(dlon) * cos(lat2)
    y = cos(lat1) * sin(lat2) - sin(lat1) * cos(lat2) * cos(dlon)
    return (degrees(atan2(x, y)) + 360) % 360


def bearing_to_color(b: float) -> str:
    """
    Map a bearing angle (degrees) to one of four directional color names.

    Quadrant boundaries are set at 45-degree diagonals:
      North (orange) : 337.5 - 360 or 0 - 22.5
      East  (green)  : 67.5 - 112.5
      South (brown)  : 157.5 - 202.5
      West  (slate)  : 247.5 - 292.5

    Bearings that fall in the diagonal zones (between quadrants) are assigned
    to the nearest principal direction to avoid ambiguity.
    """
    if b < 22.5 or b >= 337.5:
        return "orange"   # North
    elif 67.5 <= b < 112.5:
        return "green"    # East
    elif 157.5 <= b < 202.5:
        return "brown"    # South
    elif 247.5 <= b < 292.5:
        return "slate"    # West

    # Diagonal zone: pick nearest cardinal direction by angular distance.
    closest = min(
        {"orange": 0, "green": 90, "brown": 180, "slate": 270}.items(),
        key=lambda item: abs((b - item[1] + 180) % 360 - 180),
    )
    return closest[0]


# ---------------------------------------------------------------------------
# Main processing: build and color the connections table
# ---------------------------------------------------------------------------

def generate_colored_table(named_edges: list, location_coords: dict) -> None:
    """
    Build the Colored Connections Table Excel file.

    For each location row in the Connections Table, each cable is assigned a
    directional color based on the bearing measured 5 m along the cable from
    the splice point. The 5 m walk avoids the ambiguity caused by short jogs
    near enclosures, giving a stable bearing representative of the cable's
    general routing direction.

    Special override rules (applied before bearing calculation):
      - OLT cables are always colored olive green.
      - MST tap cables (single-connection FT nearest its paired SE) are red.
      - Within one location, each directional color may only appear once;
        a fallback to the next-nearest color is used if a direction repeats.
    """
    df_main = pd.read_excel(CONNECTIONS_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Colored Connections"
    ws.append(df_main.columns.tolist())   # write header row

    # PatternFill objects for each directional and special-case color.
    fills = {
        "orange": PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"),
        "brown":  PatternFill(start_color="8B4513", end_color="8B4513", fill_type="solid"),
        "green":  PatternFill(start_color="008000", end_color="008000", fill_type="solid"),
        "slate":  PatternFill(start_color="708090", end_color="708090", fill_type="solid"),
        "red":    PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),
        "olt":    PatternFill(start_color="C5D9B5", end_color="C5D9B5", fill_type="solid"),
    }

    # Build a lookup from cable name to its coordinate list.
    name_to_coords = {name: coords for name, coords in named_edges}
    red_cables     = set()   # cable names confirmed as MST (colored red)
    all_colored_rows = []    # collect (entries, conn_cols, raw_colors) for two-pass output

    # ------------------------------------------------------------------
    # Pre-scan pass: identify MST tap locations.
    #
    # An MST tap is an FT enclosure whose only cable runs to a nearby SE
    # enclosure (< 50 m away). When multiple FT taps are close to the same
    # SE (which happens frequently in dense areas), only the single nearest
    # FT is the true MST tap. The rest are normal taps colored by direction.
    #
    # This pre-scan builds a candidate list per SE and picks the winner
    # before the main coloring loop runs, so the decision is consistent.
    # ------------------------------------------------------------------
    mst_winner_locs: set[str] = set()    # location names confirmed as MST taps
    se_candidates:  dict[str, list] = {} # SE enclosure -> [(distance_m, ft_location)]

    for _, row in df_main.iterrows():
        loc        = str(row["Location"]).strip()
        latlon     = (row["Latitude"], row["Longitude"])
        conn_cols  = [c for c in row.index if str(c).startswith("Connection")]
        valid_conn = [row[c] for c in conn_cols if not pd.isna(row[c])]

        # Only single-connection locations can be MST taps.
        if len(valid_conn) != 1:
            continue

        cable   = valid_conn[0]
        matches = NEW_LOC_TOKEN_RX.findall(cable) or LEGACY_LOC_TOKEN_RX.findall(cable)
        if len(matches) != 2:
            continue    # need exactly two endpoints to know which side is SE

        end1, end2 = matches
        other = end1 if loc != end1 else end2

        # The "other" endpoint must be an SE or SD type enclosure.
        is_se_type = re.search(r"[SD]\d{3}$", other) or re.search(r"_SE_\d{3}$", other)
        if not is_se_type:
            continue

        other_coord = location_coords.get(other)
        if not other_coord:
            continue

        dist_m = geodesic(latlon, other_coord).meters
        if dist_m < 50:
            # Collect this FT tap as a candidate for the SE enclosure.
            se_candidates.setdefault(other, []).append((dist_m, loc))

    # For each SE, the closest FT tap wins the MST designation.
    for se, candidates in se_candidates.items():
        winner = min(candidates, key=lambda t: t[0])[1]
        mst_winner_locs.add(winner)

    # ------------------------------------------------------------------
    # Identify the OLT location and its primary outbound cable.
    # The OLT is the bare site token row (e.g. "RC73E") with no type suffix.
    # ------------------------------------------------------------------
    olt_pattern = re.compile(r"^[A-Z]{2}\d{2,3}E$")
    olt_loc  = None
    olt_conn = None

    for _, row in df_main.iterrows():
        loc = str(row["Location"]).strip()
        if olt_pattern.match(loc):
            olt_loc   = loc
            conn_cols = [c for c in row.index if str(c).startswith("Connection")]
            for c in conn_cols:
                if pd.notna(row[c]):
                    olt_conn = row[c].strip()
                    break
            break

    # ------------------------------------------------------------------
    # Main coloring loop: determine the direction color for each cable
    # at each location.
    # ------------------------------------------------------------------
    for _, row in df_main.iterrows():
        loc       = row["Location"]
        latlon    = (row["Latitude"], row["Longitude"])
        conn_cols = [c for c in row.index if str(c).startswith("Connection")]
        entries   = [loc, latlon[0], latlon[1]]   # start each row with location + coords
        raw       = []                            # (col, color_name, bearing) for this row
        valid_conn = [row[c] for c in conn_cols if not pd.isna(row[c])]

        for col in conn_cols:
            cable = row[col]
            if pd.isna(cable):
                entries.append("")
                continue

            # -- OLT override: color any OLT cable olive green --
            if olt_loc and is_olt_cable(cable, olt_loc):
                raw.append((col, "olt", 0))
                entries.append(cable)
                continue

            # -- Look up the cable's coordinate geometry from the KMZ --
            coords = name_to_coords.get(cable)
            if not coords:
                entries.append(cable)
                continue

            # Ensure coordinate sequence starts at the current splice location.
            if coords[-1] == latlon:
                coords = list(reversed(coords))

            # -- MST override: color confirmed MST tap cables red --
            matches = NEW_LOC_TOKEN_RX.findall(cable) or LEGACY_LOC_TOKEN_RX.findall(cable)
            if len(valid_conn) == 1 and len(matches) == 2:
                end1, end2 = matches
                other = end1 if loc != end1 else end2
                is_se_type = (
                    re.search(r"[SD]\d{3}$", other) or
                    re.search(r"_SE_\d{3}$", other)
                )
                if is_se_type and loc in mst_winner_locs:
                    raw.append((col, "red", 0))
                    red_cables.add(cable)
                    entries.append(cable)
                    continue

            # -- Walk 5 m along the cable and compute the bearing --
            # Walking a short distance away from the splice avoids the
            # erratic bearing that can result from very small initial segments.
            dist   = 0
            walked = latlon
            for i in range(len(coords) - 1):
                seg_len = geodesic(coords[i], coords[i + 1]).meters
                if dist + seg_len >= 5:
                    # Interpolate to find the exact 5 m point on this segment.
                    ratio  = (5 - dist) / seg_len
                    lat    = coords[i][0] + (coords[i + 1][0] - coords[i][0]) * ratio
                    lon    = coords[i][1] + (coords[i + 1][1] - coords[i][1]) * ratio
                    walked = (lat, lon)
                    break
                dist += seg_len

            # If the cable is shorter than 5 m, use its far endpoint.
            if walked == latlon:
                walked = coords[-1]

            b     = bearing(latlon, walked)
            color = bearing_to_color(b)
            raw.append((col, color, b))
            entries.append(cable)

        all_colored_rows.append((entries, conn_cols, raw))

    # ------------------------------------------------------------------
    # Output pass: write rows to the workbook and apply cell fills.
    #
    # Each location gets at most one cell of each directional color; if two
    # cables happen to have the same bearing category, the second is assigned
    # the next-nearest unused direction rather than a duplicate.
    # ------------------------------------------------------------------
    for entries, conn_cols, raw in all_colored_rows:
        ws.append(entries)
        used     = set()     # directional colors already assigned at this location
        assigned = {}        # col -> color_name for final fill application

        for col, colr, b in raw:
            if colr == "olt":
                assigned[col] = "olt"
                continue

            cable_val = entries[3 + conn_cols.index(col)]
            if cable_val in red_cables:
                # Cable was confirmed MST in the pre-scan; always red.
                assigned[col] = "red"
            elif colr not in used:
                # First cable with this direction at this location; use it directly.
                assigned[col] = colr
                used.add(colr)
            else:
                # Direction already used; pick the next-nearest unused direction.
                direction_angles = {"orange": 0, "green": 90, "brown": 180, "slate": 270}
                alt      = list({"orange", "green", "brown", "slate"} - used)
                if not alt:
                    alt = list(direction_angles.keys())   # all used; allow repeat
                fallback = min(
                    alt,
                    key=lambda c: abs((b - direction_angles[c] + 180) % 360 - 180),
                )
                assigned[col] = fallback
                used.add(fallback)

        # Apply fills to the connection columns (columns 4 onward in the sheet).
        loc_current = entries[0]
        for idx, col in enumerate(conn_cols):
            cable = entries[3 + idx]
            if col in assigned and assigned[col] == "olt":
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills["olt"]
            elif olt_conn and cable == olt_conn:
                # Secondary OLT cable match (cable string matches but wasn't
                # caught by is_olt_cable due to name variation).
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills["olt"]
            elif col in assigned:
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills[assigned[col]]

    wb.save(OUTPUT_PATH)
    print(f"Colored table saved to {OUTPUT_PATH}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Extract and merge KMZ cable geometry.
    fiber_segments = extract_fiber_segments()
    fiber_segments = merge_named_segments(fiber_segments)

    # Build location coordinate lookup from the Connections Table.
    df_conn = pd.read_excel(CONNECTIONS_PATH)
    location_coords = {}
    for _, row in df_conn.iterrows():
        try:
            name = str(row["Location"]).strip()
            lat  = float(row["Latitude"])
            lon  = float(row["Longitude"])
            if name and not pd.isna(lat) and not pd.isna(lon):
                location_coords[name] = (lat, lon)
        except Exception:
            continue   # skip any malformed rows silently

    generate_colored_table(fiber_segments, location_coords)