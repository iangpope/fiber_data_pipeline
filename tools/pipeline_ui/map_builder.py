"""
map_builder.py -- Build GeoJSON from pipeline step 1+2 outputs for the
checkpoint map view.

Combines:
  1. data/Connections_Table.xlsx        -- location names + GPS coords + cable names
  2. output/Colored_Connections_Table.xlsx -- cable-to-color assignments (cell fill)
  3. data/<project>.kmz                 -- cable line geometries

Outputs a single GeoJSON FeatureCollection with:
  - One MultiLineString feature per cable (all segments, neutral color)
  - One Point feature per location (markers, with per-location cable colors)
"""

from __future__ import annotations

import json
import os
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from typing import Optional

import openpyxl
import pandas as pd

# ---------------------------------------------------------------------------
# Pipeline root on sys.path for config import
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent          # tools/pipeline_ui/
_ROOT = _HERE.parent.parent                      # Fiber Data Pipeline/
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from config import COLOR


# ---------------------------------------------------------------------------
# Color mapping: openpyxl fill hex → human label + CSS hex
# ---------------------------------------------------------------------------

_FILL_TO_INFO: dict[str, dict] = {
    COLOR["NORTH"]: {"label": "North",   "css": f"#{COLOR['NORTH']}"},
    COLOR["SOUTH"]: {"label": "South",   "css": f"#{COLOR['SOUTH']}"},
    COLOR["EAST"]:  {"label": "East",    "css": f"#{COLOR['EAST']}"},
    COLOR["WEST"]:  {"label": "West",    "css": f"#{COLOR['WEST']}"},
    COLOR["OLT"]:   {"label": "OLT",     "css": f"#{COLOR['OLT']}"},
    COLOR["MST"]:   {"label": "MST",     "css": f"#{COLOR['MST']}"},
}

_DEFAULT_COLOR = {"label": "Unknown", "css": "#888888"}


def _cable_short_name(cable_name: str, loc_name: str) -> str:
    """Return compact 'other-end (size)' label for this cable at this location."""
    if '_TO_' in cable_name:
        end_a, rest = cable_name.split('_TO_', 1)
        m = re.search(r'_(\d{2,3}CT)$', rest, re.IGNORECASE)
        size  = m.group(1) if m else ''
        end_b = rest[:m.start()] if m else rest
        other = end_b if end_a == loc_name else end_a
        return f"{other} ({size})" if size else other
    m = re.match(r'^(\d{2,3}CT)\s+(\S+)\s+TO\s+(\S+)$', cable_name, re.IGNORECASE)
    if m:
        size, ea, eb = m.group(1), m.group(2), m.group(3)
        return f"{eb if ea == loc_name else ea} ({size})"
    return cable_name


def _cable_size(cable_name: str) -> str:
    """Extract the fiber-count suffix from a cable name (e.g. '48CT')."""
    m = re.search(r'(\d{2,3}CT)', cable_name, re.IGNORECASE)
    return m.group(1).upper() if m else ''


def _hex_from_fill(cell) -> Optional[str]:
    """Extract the 6-char hex color from an openpyxl cell's PatternFill."""
    try:
        fgcolor = cell.fill.fgColor
        raw = fgcolor.rgb if fgcolor.type == "rgb" else None
        if raw and len(raw) >= 6:
            return raw[-6:].upper()
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# KMZ cable geometry parser — folder-based
#
# Magellan KMZ exports organise fiber cables as:
#   fiber/fiberCable/<GUID folder name = cable name>/<Placemark LineStrings>
#
# A single cable can be split across many Placemarks (e.g. 24 segments for
# one cable).  We collect ALL segments per cable so the map renders the
# complete route rather than one fragment.
# ---------------------------------------------------------------------------

def _strip_ns(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def _get_name(el) -> str:
    n = next((c for c in el if _strip_ns(c.tag) == "name"), None)
    return (n.text or "").strip() if n is not None else ""


def _parse_kmz_cables(kmz_path: str) -> dict[str, list[list[tuple[float, float]]]]:
    """
    Return {cable_name: [[seg1_points], [seg2_points], ...]} by walking the
    fiber/fiberCable folder hierarchy in the KMZ.

    Each inner list is one LineString segment expressed as (lat, lon) tuples.
    Keeping segments separate (rather than concatenating) avoids drawing
    spurious connecting lines when segments are not geographically ordered.
    """
    cable_segments: dict[str, list] = {}

    with zipfile.ZipFile(kmz_path, "r") as kmz:
        kml_name = next((n for n in kmz.namelist() if n.endswith(".kml")), None)
        if not kml_name:
            return cable_segments
        root = ET.fromstring(kmz.read(kml_name))

    # Locate the fiber/fiberCable folder (works regardless of KML namespace).
    fiber_cable_folder = None
    for folder in root.iter():
        if _strip_ns(folder.tag) == "Folder" and _get_name(folder) == "fiberCable":
            fiber_cable_folder = folder
            break

    if fiber_cable_folder is None:
        return cable_segments

    # Each direct child Folder has the cable name; its Placemarks hold geometry.
    for guid_folder in fiber_cable_folder:
        if _strip_ns(guid_folder.tag) != "Folder":
            continue
        cable_name = _get_name(guid_folder)
        if not cable_name:
            continue

        segments = []
        for pm in guid_folder:
            if _strip_ns(pm.tag) != "Placemark":
                continue
            ls = next((c for c in pm.iter() if _strip_ns(c.tag) == "LineString"), None)
            if ls is None:
                continue
            coords_el = next((c for c in ls.iter() if _strip_ns(c.tag) == "coordinates"), None)
            if coords_el is None or not coords_el.text:
                continue
            points = []
            for tok in coords_el.text.split():
                parts = tok.split(",")
                if len(parts) >= 2:
                    try:
                        points.append((float(parts[1]), float(parts[0])))  # lat, lon
                    except ValueError:
                        pass
            if points:
                segments.append(points)

        if segments:
            cable_segments[cable_name] = segments

    return cable_segments


# ---------------------------------------------------------------------------
# Main public function
# ---------------------------------------------------------------------------

def build_geojson(job_dir: str) -> dict:
    """
    Build a GeoJSON FeatureCollection from the outputs of pipeline steps 1 and 2.

    Parameters
    ----------
    job_dir : str
        Path to the job's temp directory (contains data/ and output/ sub-dirs).

    Returns
    -------
    dict -- GeoJSON FeatureCollection ready for json.dumps().
            Returns an empty FeatureCollection on any error.
    """
    data_dir   = Path(job_dir) / "data"
    output_dir = Path(job_dir) / "output"

    conn_table_path    = data_dir   / "Connections_Table.xlsx"
    colored_table_path = output_dir / "Colored_Connections_Table.xlsx"

    if not conn_table_path.exists() or not colored_table_path.exists():
        return {"type": "FeatureCollection", "features": []}

    # Find KMZ file.
    kmz_path = next(
        (str(data_dir / f) for f in os.listdir(data_dir) if f.lower().endswith(".kmz")),
        None,
    )

    # Load direction confidence data written by step 2 (absent on CLI-only runs).
    cable_confidence: dict[str, dict] = {}
    conf_summary: dict = {}
    _conf_path = output_dir / "direction_confidence.json"
    if _conf_path.exists():
        with open(_conf_path) as _f:
            _cdata = json.load(_f)
            cable_confidence = _cdata.get("cables", {})
            conf_summary     = _cdata.get("summary", {})

    # -----------------------------------------------------------------------
    # 1. Read cable → color from the Colored Connections Table.
    # -----------------------------------------------------------------------
    cable_info: dict[str, dict] = {}   # cable_name → {label, css}

    wb_colored = openpyxl.load_workbook(colored_table_path)
    ws = wb_colored.active

    max_col = ws.max_column
    for row in ws.iter_rows(min_row=2, max_col=max_col):
        for cell in row[3:]:   # skip Location, Lat, Lon columns
            val = str(cell.value or "").strip()
            if not val:
                continue
            hex6 = _hex_from_fill(cell)
            info = _FILL_TO_INFO.get(hex6, _DEFAULT_COLOR) if hex6 else _DEFAULT_COLOR
            cable_info[val] = info

    # -----------------------------------------------------------------------
    # 2. Read location nodes from Connections Table.
    # -----------------------------------------------------------------------
    df = pd.read_excel(conn_table_path)
    location_data: dict[str, dict] = {}  # name → {lat, lon, cables:[...]}

    for _, row in df.iterrows():
        loc = str(row.get("Location", "")).strip()
        try:
            lat = float(row.get("Latitude", 0))
            lon = float(row.get("Longitude", 0))
        except (ValueError, TypeError):
            continue
        if not loc or lat == 0:
            continue

        cables = []
        for col in df.columns:
            if str(col).startswith("Connection"):
                val = str(row.get(col, "")).strip()
                if val and val.lower() not in ("nan", "none", ""):
                    info = cable_info.get(val, _DEFAULT_COLOR)
                    cables.append({
                        "name":       val,
                        "color":      info["css"],
                        "label":      info["label"],
                        "short_name": _cable_short_name(val, loc),
                        "size":       _cable_size(val),
                    })

        location_data[loc] = {"lat": lat, "lon": lon, "cables": cables}

    # -----------------------------------------------------------------------
    # 3. Parse KMZ cable geometries (all segments per cable).
    # -----------------------------------------------------------------------
    kmz_cables = _parse_kmz_cables(kmz_path) if kmz_path else {}

    # -----------------------------------------------------------------------
    # 4. Build GeoJSON features.
    # -----------------------------------------------------------------------
    features = []

    # --- Cable MultiLineString features ---
    # Color is stored in properties for node-click highlighting and override
    # tracking, but lines are drawn neutral on the map until a node is clicked.
    seen_cables: set[str] = set()
    for cable, info in cable_info.items():
        if cable in seen_cables:
            continue
        seen_cables.add(cable)

        segments = kmz_cables.get(cable)
        if not segments:
            continue

        # GeoJSON coordinates are [lon, lat]
        coords_geojson = [
            [[lon, lat] for lat, lon in seg]
            for seg in segments
        ]

        _conf = cable_confidence.get(cable, {})
        features.append({
            "type": "Feature",
            "geometry": {
                "type":        "MultiLineString",
                "coordinates": coords_geojson,
            },
            "properties": {
                "cable":      cable,
                "color":      info["css"],
                "direction":  info["label"],
                "bearing":    _conf.get("bearing"),
                "confidence": _conf.get("confidence", "ok"),
            },
        })

    _missing_cables = {c for c, d in cable_confidence.items() if d.get("confidence") == "missing"}

    # --- Location point features ---
    for loc, data in location_data.items():
        missing_here = [c["name"] for c in data["cables"] if c["name"] in _missing_cables]
        features.append({
            "type": "Feature",
            "geometry": {
                "type":        "Point",
                "coordinates": [data["lon"], data["lat"]],
            },
            "properties": {
                "name":           loc,
                "cables":         data["cables"],
                "missing_cables": missing_here,
            },
        })

    return {"type": "FeatureCollection", "features": features, "summary": conf_summary}
