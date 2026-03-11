"""
1_coords_from_kmz.py -- Extract splice location coordinates from a KMZ file
and build the Connections Table used by the rest of the pipeline.

Reads:
    data/<project>.kmz   -- KMZ exported from GIS containing placemarks for
                            every splice enclosure with GPS coordinates
    data/<cut_sheet>.xlsx -- The cut sheet workbook exported from the design
                             tool (any .xlsx in the data folder)

Writes:
    data/cut_sheet.xlsx          -- Normalized copy of the cut sheet
    data/Connections_Table.xlsx  -- One row per splice location, with its
                                    GPS coordinates and all cable connections
                                    listed in separate columns
"""

import os
import shutil
import pandas as pd
import zipfile
from xml.etree import ElementTree as ET


# ---------------------------------------------------------------------------
# Directory setup
# ---------------------------------------------------------------------------
DATA_DIR   = "data"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# Step 1: Locate and normalize the cut sheet.
#
# The cut sheet may have any filename when it arrives in the data folder.
# We copy it to a fixed name (cut_sheet.xlsx) so every downstream script
# can reference a known path without searching the folder again.
#
# Several auxiliary xlsx files also live in data/ (HAF report, Tap Report
# template, Connections Table from a prior run, etc.).  These are excluded
# by name so that only the genuine GIS cut sheet is selected.
# ---------------------------------------------------------------------------
_EXCLUDE = {
    "cut_sheet.xlsx",
    "connections_table.xlsx",
}

def _is_auxiliary(filename: str) -> bool:
    """Return True if filename looks like an auxiliary data file, not the cut sheet."""
    lower = filename.lower()
    return (
        lower in _EXCLUDE
        or "haf" in lower
        or "template" in lower
        or "tap report" in lower
    )

xlsx_files = [
    f for f in os.listdir(DATA_DIR)
    if f.lower().endswith(".xlsx") and not _is_auxiliary(f)
]
if not xlsx_files:
    raise FileNotFoundError(
        "No cut sheet found in data/. Place the GIS cut sheet xlsx there and retry."
    )

cut_sheet_original = os.path.join(DATA_DIR, sorted(xlsx_files)[0])
cut_sheet_path     = os.path.join(DATA_DIR, "cut_sheet.xlsx")

# Only copy if the file is not already named correctly (avoids a redundant I/O).
if cut_sheet_original != cut_sheet_path:
    shutil.copy(cut_sheet_original, cut_sheet_path)


# ---------------------------------------------------------------------------
# Step 2: Locate the KMZ file.
#
# One KMZ per project is expected in the data folder.
# ---------------------------------------------------------------------------
kmz_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".kmz")]
if not kmz_files:
    raise FileNotFoundError("No .kmz file found in data folder.")
kmz_path = os.path.join(DATA_DIR, kmz_files[0])


# ---------------------------------------------------------------------------
# Step 3: Parse GPS coordinates from the KMZ.
#
# A KMZ is a ZIP archive containing at least one KML file. Each splice
# location appears as a <Placemark> element with a <name> (location ID) and
# a <coordinates> element in lon,lat[,alt] order. We build a dict that maps
# location name -> (lat, lon) for fast lookup when building the connections table.
# ---------------------------------------------------------------------------
location_coords = {}

with zipfile.ZipFile(kmz_path, "r") as kmz:
    # Find the KML file inside the archive (usually named "doc.kml").
    kml_filename = next(
        (name for name in kmz.namelist() if name.endswith(".kml")), None
    )
    if not kml_filename:
        raise FileNotFoundError("No .kml file found inside the KMZ archive.")

    kml_content = kmz.read(kml_filename)
    root = ET.fromstring(kml_content)

    # KML uses an XML namespace; register it for .find() calls.
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    for pm in root.findall(".//kml:Placemark", ns):
        name_elem   = pm.find("kml:name", ns)
        coords_elem = pm.find(".//kml:coordinates", ns)

        # Skip placemarks that are missing either a name or coordinates.
        if name_elem is None or coords_elem is None:
            continue

        name        = name_elem.text.strip()
        coords_text = coords_elem.text.strip()
        if not coords_text:
            continue

        # KML coordinates are space-separated; each value is "lon,lat[,alt]".
        # For point placemarks (single splice locations) there is only one
        # coordinate set. For line placemarks we use the first point only.
        coords_list = coords_text.replace("\n", " ").split()
        if not coords_list:
            continue

        lon_lat_alt = coords_list[0].split(",")   # split "lon,lat,alt"
        if len(lon_lat_alt) < 2:
            continue

        lon = float(lon_lat_alt[0])
        lat = float(lon_lat_alt[1])
        location_coords[name] = (lat, lon)


# ---------------------------------------------------------------------------
# Step 4: Build the Connections Table from the cut sheet.
#
# Each sheet in the cut sheet workbook corresponds to one splice location.
# We scan column B (index 1) for rows that contain the word "TO", which
# indicates a cable connection (e.g. "RC73E_FT_001 TO RC73E_SE_001 48CT").
# Each unique cable string found becomes a Connection column in the output.
# ---------------------------------------------------------------------------
xls = pd.ExcelFile(cut_sheet_path)
connections_data = []

for sheet_name in xls.sheet_names:
    # Skip administrative/reference sheets that are not splice locations.
    if sheet_name.lower() in ["connections", "summary", "index"]:
        continue

    df_sheet = xls.parse(sheet_name, header=None)

    try:
        # Only process sheets that have a corresponding GPS coordinate entry.
        if sheet_name not in location_coords:
            continue

        lat, lon = location_coords[sheet_name]

        # Find rows in column B that contain the word "TO" (cable connections).
        mask        = df_sheet.iloc[:, 1].astype(str).str.contains("TO", case=False, na=False)
        connections = df_sheet.loc[mask, 1].dropna().unique()

        # Build one dict per location; connections go into numbered columns.
        entry = {"Location": sheet_name, "Latitude": lat, "Longitude": lon}
        for i, conn in enumerate(connections, start=1):
            entry[f"Connection {i}"] = str(conn).strip()

        connections_data.append(entry)

    except Exception as e:
        print(f"Skipping sheet {sheet_name} due to error: {e}")
        continue


# ---------------------------------------------------------------------------
# Step 5: Write the Connections Table to Excel.
# ---------------------------------------------------------------------------
df_connections = pd.DataFrame(connections_data)
df_connections = df_connections.sort_values(by="Location")   # alphabetical order

output_path = os.path.join(DATA_DIR, "Connections_Table.xlsx")
df_connections.to_excel(output_path, index=False)
print(f"Wrote {len(df_connections)} locations to {output_path}")
