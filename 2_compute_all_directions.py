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

# Support both new (MS90E_FT_001 / MS90E_SE_001) and legacy (MIC...) location tokens
# No \b boundaries — underscores between tokens prevent \b from firing
NEW_LOC_TOKEN_RX = re.compile(r"[A-Z0-9]+E_(?:FT|SE)_\d{3}")
LEGACY_LOC_TOKEN_RX = re.compile(r"\bMIC\w+\b")


OLT_TOKEN_RX = re.compile(r"^[A-Z]{2}\d{2,3}E$")
FIBER_SUFFIX_RX = re.compile(r"_(\d{2,3}CT)\b", flags=re.IGNORECASE)

def parse_endpoints(cable_name: str):
    """Return (endA, endB, fiber) from either:
      - '48CT LOC_A TO LOC_B'
      - 'LOC_A_TO_LOC_B_48CT' (RC077E style)
    """
    t = str(cable_name or "").strip()
    if not t:
        return None, None, None

    if "_TO_" in t:
        a, b = t.split("_TO_", 1)
        fiber = None
        m = re.search(r"_(\d{2,3}CT)\b", b, flags=re.IGNORECASE)
        if m:
            fiber = m.group(1).upper()
            b = re.sub(r"_(\d{2,3}CT)\b", "", b, flags=re.IGNORECASE)
        return a.strip(), b.strip(), fiber

    m = re.match(r"^\s*(\d{2,3}CT)\s+(\S+)\s+TO\s+(\S+)\s*$", t, flags=re.IGNORECASE)
    if m:
        fiber = m.group(1).upper()
        return m.group(2).strip(), m.group(3).strip(), fiber

    return None, None, None

def is_olt_cable(cable_name: str, olt_token: str) -> bool:
    a, b, _ = parse_endpoints(cable_name)
    if not olt_token:
        return False
    return (a == olt_token) or (b == olt_token)

# Constants
DATA_DIR = "data"
CONNECTIONS_PATH = os.path.join(DATA_DIR, "Connections_Table.xlsx")
OUTPUT_PATH = "output/Colored_Connections_Table.xlsx"

# Auto-detect KMZ file in data/
def detect_kmz_file():
    for f in os.listdir(DATA_DIR):
        if f.lower().endswith(".kmz"):
            return os.path.join(DATA_DIR, f)
    raise FileNotFoundError("No .kmz file found in 'data' folder.")

KMZ_PATH = detect_kmz_file()

# Namespace-agnostic KML parser
def extract_fiber_segments():
    def strip_ns(tag):
        return tag.split('}')[-1] if '}' in tag else tag

    with zipfile.ZipFile(KMZ_PATH, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.endswith(".kml")), None)
        root = ET.fromstring(zf.read(kml_file))

    def recursive_find(el):
        segments = []
        for child in el.iter():
            tag = strip_ns(child.tag.lower())
            if tag == "placemark":
                coords_elem = child.find(".//{*}LineString/{*}coordinates")
                if coords_elem is not None:
                    coords = coords_elem.text.strip().split()
                    latlons = [(float(c.split(",")[1]), float(c.split(",")[0])) for c in coords if "," in c]
                    name_elem = next((c for c in child if strip_ns(c.tag) == "name"), None)
                    name = name_elem.text.strip() if (name_elem is not None and name_elem.text) else None
                    segments.append((name if name else None, latlons))
        return segments

    segments = recursive_find(root)
    print(f"📦 Extracted {len(segments)} line segments from KMZ")
    return segments

# Merge segment coordinates by unique cable name
def merge_named_segments(segments):
    from collections import defaultdict, deque
    grouped = defaultdict(list)
    for name, coords in segments:
        if name:
            grouped[name].append(coords)
    merged = []
    for name, paths in grouped.items():
        if len(paths) == 1:
            merged.append((name, paths[0]))
            continue

        chain = deque(paths[0])
        remaining = list(paths[1:])
        while remaining:
            progress = False
            still_remaining = []
            for p in remaining:
                if chain[-1] == p[0]:             # append forward
                    chain.extend(p[1:])
                    progress = True
                elif chain[-1] == p[-1]:          # append reversed
                    chain.extend(reversed(p[:-1]))
                    progress = True
                elif chain[0] == p[-1]:           # prepend forward
                    chain.extendleft(reversed(p[:-1]))
                    progress = True
                elif chain[0] == p[0]:            # prepend reversed
                    chain.extendleft(p[1:])
                    progress = True
                else:
                    still_remaining.append(p)
            if not progress:
                # Segments don't connect geometrically; append to avoid data loss
                for p in still_remaining:
                    chain.extend(p)
                break
            remaining = still_remaining
        merged.append((name, list(chain)))
    return merged

# Compute bearing and color
def bearing(p1, p2):
    lat1, lon1 = radians(p1[0]), radians(p1[1])
    lat2, lon2 = radians(p2[0]), radians(p2[1])
    dlon = lon2 - lon1
    x = sin(dlon) * cos(lat2)
    y = cos(lat1) * sin(lat2) - sin(lat1) * cos(lat2) * cos(dlon)
    return (degrees(atan2(x, y)) + 360) % 360

def bearing_to_color(b):
    if b < 22.5 or b >= 337.5:
        return "orange"
    elif 67.5 <= b < 112.5:
        return "green"
    elif 157.5 <= b < 202.5:
        return "brown"
    elif 247.5 <= b < 292.5:
        return "slate"
    closest = min({"orange": 0, "green": 90, "brown": 180, "slate": 270}.items(),
                  key=lambda item: abs((b - item[1] + 180) % 360 - 180))
    return closest[0]

# Output to Excel
def generate_colored_table(named_edges, location_coords):
    df_main = pd.read_excel(CONNECTIONS_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Colored Connections"
    ws.append(df_main.columns.tolist())

    fills = {
        "orange": PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"),
        "brown": PatternFill(start_color="8B4513", end_color="8B4513", fill_type="solid"),  # improved brown
        "green": PatternFill(start_color="008000", end_color="008000", fill_type="solid"),
        "slate": PatternFill(start_color="708090", end_color="708090", fill_type="solid"),
        "red": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),
        "olt": PatternFill(start_color="C5D9B5", end_color="C5D9B5", fill_type="solid")  # olive green
    }

    name_to_coords = {name: coords for name, coords in named_edges}
    red_cables = set()
    all_colored_rows = []

    # -----------------------------------------------------------------------
    # Pre-scan: for each SE enclosure, find the single nearest FT tap.
    # Only that one winner will be colored red (MST).
    # This prevents over-identification when multiple FT taps are near the
    # same SE (only the closest one is the true MST tap).
    # -----------------------------------------------------------------------
    mst_winner_locs: set[str] = set()   # location names that are confirmed MST taps
    se_candidates: dict[str, list] = {} # SE_name -> [(dist, ft_loc)]

    for _, row in df_main.iterrows():
        loc = str(row["Location"]).strip()
        latlon = (row["Latitude"], row["Longitude"])
        conn_cols = [c for c in row.index if str(c).startswith("Connection")]
        valid_conn = [row[c] for c in conn_cols if not pd.isna(row[c])]

        if len(valid_conn) != 1:
            continue
        cable = valid_conn[0]
        matches = NEW_LOC_TOKEN_RX.findall(cable) or LEGACY_LOC_TOKEN_RX.findall(cable)
        if len(matches) != 2:
            continue
        end1, end2 = matches
        other = end1 if loc != end1 else end2
        is_se_type = re.search(r"[SD]\d{3}$", other) or re.search(r"_SE_\d{3}$", other)
        if not is_se_type:
            continue
        other_coord = location_coords.get(other)
        if not other_coord:
            continue
        dist_m = geodesic(latlon, other_coord).meters
        if dist_m < 50:
            se_candidates.setdefault(other, []).append((dist_m, loc))

    for se, candidates in se_candidates.items():
        # Winner = closest FT tap to this SE
        winner = min(candidates, key=lambda t: t[0])[1]
        mst_winner_locs.add(winner)

    # Detect OLT row and its connection
    olt_pattern = re.compile(r"^[A-Z]{2}\d{2,3}E$")
    olt_loc = None
    olt_conn = None
    for _, row in df_main.iterrows():
        loc = str(row["Location"]).strip()
        if olt_pattern.match(loc):
            olt_loc = loc
            conn_cols = [c for c in row.index if str(c).startswith("Connection")]
            for c in conn_cols:
                if pd.notna(row[c]):
                    olt_conn = row[c].strip()
                    break
            break

    for _, row in df_main.iterrows():
        loc = row["Location"]
        latlon = (row["Latitude"], row["Longitude"])
        conn_cols = [c for c in row.index if str(c).startswith("Connection")]
        entries = [loc, latlon[0], latlon[1]]
        raw = []
        valid_conn = [row[c] for c in conn_cols if not pd.isna(row[c])]

        for col in conn_cols:
            cable = row[col]
            if pd.isna(cable):
                entries.append("")
                continue


            # OLT override: if one endpoint is the bare OLT token (e.g., RC077E), force pale-green
            if olt_loc and is_olt_cable(cable, olt_loc):
                raw.append((col, "olt", 0))
                entries.append(cable)
                continue

            coords = name_to_coords.get(cable)
            if not coords:
                entries.append(cable)
                continue

            if coords[-1] == latlon:
                coords = list(reversed(coords))

            matches = NEW_LOC_TOKEN_RX.findall(cable) or LEGACY_LOC_TOKEN_RX.findall(cable)
            if len(valid_conn) == 1 and len(matches) == 2:
                end1, end2 = matches
                other = end1 if loc != end1 else end2
                is_se_type = re.search(r"[SD]\d{3}$", other) or re.search(r"_SE_\d{3}$", other)
                if is_se_type and loc in mst_winner_locs:
                    raw.append((col, "red", 0))
                    red_cables.add(cable)
                    entries.append(cable)
                    continue

            dist = 0
            walked = latlon
            for i in range(len(coords) - 1):
                seg_len = geodesic(coords[i], coords[i + 1]).meters
                if dist + seg_len >= 5:
                    ratio = (5 - dist) / seg_len
                    lat = coords[i][0] + (coords[i + 1][0] - coords[i][0]) * ratio
                    lon = coords[i][1] + (coords[i + 1][1] - coords[i][1]) * ratio
                    walked = (lat, lon)
                    break
                dist += seg_len

            # If cable shorter than 5m, use its far end for direction
            if walked == latlon:
                walked = coords[-1]

            b = bearing(latlon, walked)
            color = bearing_to_color(b)
            raw.append((col, color, b))
            entries.append(cable)

        all_colored_rows.append((entries, conn_cols, raw))

    for entries, conn_cols, raw in all_colored_rows:
        ws.append(entries)
        used = set()
        assigned = {}
        for col, colr, b in raw:
            if colr == "olt":
                assigned[col] = "olt"
                continue
            if entries[3 + conn_cols.index(col)] in red_cables:
                assigned[col] = "red"
            elif colr not in used:
                assigned[col] = colr
                used.add(colr)
            else:
                direction_angles = {"orange": 0, "green": 90, "brown": 180, "slate": 270}
                alt = list({"orange", "green", "brown", "slate"} - used)
                if not alt:
                    alt = list(direction_angles.keys())
                fallback = min(alt, key=lambda c: abs((b - direction_angles[c] + 180) % 360 - 180))
                assigned[col] = fallback
                used.add(fallback)

        loc_current = entries[0]
        for idx, col in enumerate(conn_cols):
            cable = entries[3 + idx]
            # Prefer explicit assignment (incl. OLT override)
            if col in assigned and assigned[col] == "olt":
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills["olt"]
            elif olt_conn and cable == olt_conn:
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills["olt"]
            elif col in assigned:
                ws.cell(row=ws.max_row, column=4 + idx).fill = fills[assigned[col]]

    wb.save(OUTPUT_PATH)
    print(f"✅ Colored table saved to {OUTPUT_PATH}")

if __name__ == "__main__":
    fiber_segments = extract_fiber_segments()
    fiber_segments = merge_named_segments(fiber_segments)

    df_conn = pd.read_excel(CONNECTIONS_PATH)
    location_coords = {}
    for _, row in df_conn.iterrows():
        try:
            name = str(row["Location"]).strip()
            lat = float(row["Latitude"])
            lon = float(row["Longitude"])
            if name and not pd.isna(lat) and not pd.isna(lon):
                location_coords[name] = (lat, lon)
        except:
            continue

    named_edges = fiber_segments
    generate_colored_table(named_edges, location_coords)