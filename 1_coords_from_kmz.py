import os
import shutil
import pandas as pd
import zipfile
from xml.etree import ElementTree as ET

# Constants
DATA_DIR = "data"
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Auto-detect and rename cut sheet (Excel file)
xlsx_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".xlsx")]
if not xlsx_files:
    raise FileNotFoundError("No .xlsx cut sheet found in data folder.")
cut_sheet_original = os.path.join(DATA_DIR, xlsx_files[0])
cut_sheet_path = os.path.join(DATA_DIR, "cut_sheet.xlsx")
if cut_sheet_original != cut_sheet_path:
    shutil.copy(cut_sheet_original, cut_sheet_path)

# Auto-detect KMZ file
kmz_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".kmz")]
if not kmz_files:
    raise FileNotFoundError("No .kmz file found in data folder.")
kmz_path = os.path.join(DATA_DIR, kmz_files[0])

# Extract coordinates from the KMZ (assumes a KML inside the KMZ with placemarks)
location_coords = {}
with zipfile.ZipFile(kmz_path, 'r') as kmz:
    # Locate the KML file inside the KMZ (typically "doc.kml")
    kml_filename = None
    for name in kmz.namelist():
        if name.endswith(".kml"):
            kml_filename = name
            break
    if not kml_filename:
        raise FileNotFoundError("No .kml file found inside the KMZ archive.")
    # Parse the KML content
    kml_content = kmz.read(kml_filename)
    root = ET.fromstring(kml_content)
    # Define KML namespace (if present)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    # Find all Placemark entries and extract their names and coordinates
    for pm in root.findall(".//kml:Placemark", ns):
        name_elem = pm.find("kml:name", ns)
        coords_elem = pm.find(".//kml:coordinates", ns)
        if name_elem is None or coords_elem is None:
            continue
        name = name_elem.text.strip()
        coords_text = coords_elem.text.strip()
        if not coords_text:
            continue
        # If multiple coordinates (e.g., a path), use the first set
        coords_list = coords_text.replace("\n", " ").split()
        if len(coords_list) == 0:
            continue
        lon_lat_alt = coords_list[0].split(",")  # format: lon,lat[,alt]
        if len(lon_lat_alt) < 2:
            continue
        lon = float(lon_lat_alt[0])
        lat = float(lon_lat_alt[1])
        location_coords[name] = (lat, lon)

# Build the connections table using the Excel cut sheet and KMZ coordinates
xls = pd.ExcelFile(cut_sheet_path)
connections_data = []
for sheet_name in xls.sheet_names:
    # Skip sheets that are not actual location data
    if sheet_name.lower() in ["connections", "summary", "index"]:
        continue
    df_sheet = xls.parse(sheet_name, header=None)
    try:
        # Use coordinates from KMZ for this location (sheet)
        if sheet_name not in location_coords:
            # Skip if no coordinate found for this sheet name
            continue
        lat, lon = location_coords[sheet_name]
        # Find all entries in the second column containing "TO" (case-insensitive)
        mask = df_sheet.iloc[:, 1].astype(str).str.contains("TO", case=False, na=False)
        connections = df_sheet.loc[mask, 1].dropna().unique()
        # Prepare the entry for this location
        entry = {"Location": sheet_name, "Latitude": lat, "Longitude": lon}
        for i, conn in enumerate(connections, start=1):
            entry[f"Connection {i}"] = str(conn).strip()
        connections_data.append(entry)
    except Exception as e:
        print(f"Skipping sheet {sheet_name} due to error: {e}")
        continue

# Create DataFrame and sort by Location
df_connections = pd.DataFrame(connections_data)
df_connections = df_connections.sort_values(by="Location")

# Save the connections table to Excel
output_path = os.path.join(DATA_DIR, "Connections_Table.xlsx")
df_connections.to_excel(output_path, index=False)
print(df_connections.head())
