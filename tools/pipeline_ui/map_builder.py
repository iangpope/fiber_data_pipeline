"""
map_builder.py -- Build GeoJSON from pipeline step 1+2 outputs for the
checkpoint map view.

Combines:
  1. data/Connections_Table.xlsx        -- location names + GPS coords + cable names
  2. output/Colored_Connections_Table.xlsx -- cable-to-color assignments (cell fill)
  3. data/<project>.kmz                 -- cable line geometries

Outputs a single GeoJSON FeatureCollection with:
  - One LineString feature per cable (colored by direction)
  - One Point feature per location (markers, with popup data)
"""

from __future__ import annotations

import json
import os
import re
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
# KMZ line geometry parser
# ---------------------------------------------------------------------------

def _parse_kmz_lines(kmz_path: str) -> dict[str, list[tuple[float, float]]]:
    """
    Return {cable_name: [(lat, lon), ...]} for every LineString Placemark
    in the KMZ file.
    """
    name_to_coords: dict[str, list] = {}

    with zipfile.ZipFile(kmz_path, "r") as kmz:
        kml_name = next((n for n in kmz.namelist() if n.endswith(".kml")), None)
        if not kml_name:
            return name_to_coords
        root = ET.fromstring(kmz.read(kml_name))

    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    for pm in root.findall(".//kml:Placemark", ns):
        name_el  = pm.find("kml:name", ns)
        coords_el = pm.find(".//kml:LineString/kml:coordinates", ns)
        if name_el is None or coords_el is None:
            continue

        cable = (name_el.text or "").strip()
        if not cable:
            continue

        points = []
        for tok in (coords_el.text or "").split():
            parts = tok.split(",")
            if len(parts) >= 2:
                try:
                    lon, lat = float(parts[0]), float(parts[1])
                    points.append((lat, lon))
                except ValueError:
                    pass

        if points:
            name_to_coords[cable] = points

    return name_to_coords


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

    # Load direction confidence data written by step 2 (absent on first run or CLI-only runs).
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
    #    Each cable name is in a cell; its background fill is the direction color.
    # -----------------------------------------------------------------------
    cable_info: dict[str, dict] = {}   # cable_name → {label, css}

    wb_colored = openpyxl.load_workbook(colored_table_path)
    ws = wb_colored.active

    # Header row: Location, Latitude, Longitude, Connection 1, Connection 2, ...
    # Data starts at row 2.
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
                    cables.append({"name": val, "color": info["css"], "label": info["label"]})

        location_data[loc] = {"lat": lat, "lon": lon, "cables": cables}

    # -----------------------------------------------------------------------
    # 3. Parse KMZ line geometries.
    # -----------------------------------------------------------------------
    kmz_lines = _parse_kmz_lines(kmz_path) if kmz_path else {}

    # -----------------------------------------------------------------------
    # 4. Build GeoJSON features.
    # -----------------------------------------------------------------------
    features = []

    # --- Cable line features ---
    seen_cables: set[str] = set()
    for cable, info in cable_info.items():
        if cable in seen_cables:
            continue
        seen_cables.add(cable)

        coords_latlon = kmz_lines.get(cable)
        if not coords_latlon:
            continue

        # GeoJSON coordinates are [lon, lat]
        coords_geojson = [[lon, lat] for lat, lon in coords_latlon]

        _conf = cable_confidence.get(cable, {})
        features.append({
            "type": "Feature",
            "geometry": {
                "type": "LineString",
                "coordinates": coords_geojson,
            },
            "properties": {
                "cable":      cable,
                "color":      info["css"],
                "direction":  info["label"],
                "bearing":    _conf.get("bearing"),
                "confidence": _conf.get("confidence", "ok"),
                "weight":     3,
            },
        })

    _missing_cables = {c for c, d in cable_confidence.items() if d.get("confidence") == "missing"}

    # --- Location point features ---
    for loc, data in location_data.items():
        missing_here = [c["name"] for c in data["cables"] if c["name"] in _missing_cables]
        features.append({
            "type": "Feature",
            "geometry": {
                "type": "Point",
                "coordinates": [data["lon"], data["lat"]],
            },
            "properties": {
                "name":           loc,
                "cables":         data["cables"],
                "missing_cables": missing_here,
            },
        })

    return {"type": "FeatureCollection", "features": features, "summary": conf_summary}
