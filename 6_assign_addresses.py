"""
6_assign_addresses.py -- Match each splice location to a nearby street address.

Reads the address placemarks stored in the KMZ (each address is a KML Folder
whose name is the street address string and whose first Placemark gives the
GPS pin location). For each sheet in the processed workbook, the nearest
address (by Haversine distance) is written into cell B5.

Reads:  data/<project>.kmz
        data/Connections_Table.xlsx        (used for location -> GPS lookup)
        output/Combined_Formatted_Output_processed.xlsx
Writes: output/Combined_Formatted_Output_with_Addresses.xlsx
"""

import os
import zipfile
from xml.etree import ElementTree as ET
import pandas as pd
from math import radians, cos, sin, sqrt, atan2
from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# Directory and file paths
# ---------------------------------------------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Detect the KMZ file in the data folder; exactly one is expected.
kmz_files = [f for f in os.listdir(DATA_DIR) if f.lower().endswith(".kmz")]
if len(kmz_files) != 1:
    raise FileNotFoundError(
        f"Expected 1 KMZ file in {DATA_DIR}, found {len(kmz_files)}"
    )
KMZ_PATH = os.path.join(DATA_DIR, kmz_files[0])
print(f"Using KMZ file: {kmz_files[0]}")

CONNECTIONS_PATH    = os.path.join(DATA_DIR,   "Connections_Table.xlsx")
PROCESSED_XLSX_PATH = os.path.join(OUTPUT_DIR, "Combined_Formatted_Output_processed.xlsx")
FINAL_OUTPUT_PATH   = os.path.join(OUTPUT_DIR, "Combined_Formatted_Output_with_Addresses.xlsx")


# ---------------------------------------------------------------------------
# Distance calculation
# ---------------------------------------------------------------------------

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """
    Return the great-circle distance in metres between two GPS coordinates.

    Uses the Haversine formula, which is accurate over the short distances
    (<1 km) typical within a single FTTH project area.
    """
    R     = 6_371_000   # Earth's mean radius in metres
    phi1, phi2 = radians(lat1), radians(lat2)
    d_phi      = radians(lat2 - lat1)
    d_lambda   = radians(lon2 - lon1)
    a = (
        sin(d_phi / 2) ** 2
        + cos(phi1) * cos(phi2) * sin(d_lambda / 2) ** 2
    )
    return 2 * R * atan2(sqrt(a), sqrt(1 - a))


# ---------------------------------------------------------------------------
# KMZ coordinate parsing
# ---------------------------------------------------------------------------

def safe_extract_coordinates(text: str):
    """
    Parse the first coordinate pair from a KML <coordinates> text block.

    KML stores coordinates as "lon,lat[,alt]" separated by whitespace. This
    function returns (lat, lon) for the first coordinate in the block, or
    (None, None) if the text cannot be parsed.
    """
    try:
        coords    = text.strip().replace("\n", " ").split()
        lon, lat  = map(float, coords[0].split(",")[:2])
        return lat, lon
    except Exception:
        return None, None


def extract_addresses_from_kmz(kmz_path: str) -> list:
    """
    Extract address placemarks from the KMZ and return them as a list of dicts.

    In the KMZ, address data is stored in KML <Folder> elements whose <name>
    contains a comma (e.g. "123 Main St, Anytown"). Each folder contains at
    least one <Placemark> with a <coordinates> element that gives the GPS pin
    location for that address.

    Folders that do not contain "MI" in their name or lack a comma (which
    filters out non-address folders like layer groups) are skipped.

    Returns a list of {"Address": str, "Latitude": float, "Longitude": float}.
    """
    addresses = []

    with zipfile.ZipFile(kmz_path, "r") as kmz:
        kml_file = next(
            (f for f in kmz.namelist() if f.endswith(".kml")), None
        )
        if not kml_file:
            raise FileNotFoundError("No KML file found inside the KMZ.")
        kml_content = kmz.read(kml_file)
        root = ET.fromstring(kml_content)

    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    for folder in root.findall(".//kml:Folder", ns):
        folder_name_elem = folder.find("kml:name", ns)
        if folder_name_elem is None:
            continue

        folder_name = folder_name_elem.text.strip()

        # Filter to address folders only: must contain "MI" and a comma.
        if "MI" not in folder_name or "," not in folder_name:
            continue

        # Extract the GPS pin for this address from the first placemark in the folder.
        placemark  = folder.find(".//kml:Placemark", ns)
        coords_elem = (
            placemark.find(".//kml:coordinates", ns)
            if placemark is not None
            else None
        )
        if coords_elem is None:
            continue

        lat, lon = safe_extract_coordinates(coords_elem.text)
        if lat is not None and lon is not None:
            addresses.append({
                "Address":   folder_name,
                "Latitude":  lat,
                "Longitude": lon,
            })

    print(f"Extracted {len(addresses)} address placemarks from KMZ.")
    return addresses


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> None:
    try:
        # Load the GPS coordinates for each splice location from the
        # Connections Table built in step 1.
        print("Loading connection coordinates...")
        df_conn = pd.read_excel(CONNECTIONS_PATH)
        location_coords = {
            row["Location"]: (row["Latitude"], row["Longitude"])
            for _, row in df_conn.iterrows()
        }
        print(f"Loaded {len(location_coords)} locations.")

        # Extract street addresses and their GPS pins from the KMZ.
        addresses = extract_addresses_from_kmz(KMZ_PATH)

        # Open the processed workbook from step 5.
        print("Loading output workbook...")
        wb = load_workbook(PROCESSED_XLSX_PATH)

        updated_count = 0
        for sheet_name in wb.sheetnames:
            # Skip sheets that have no GPS coordinate in the connections table.
            if sheet_name not in location_coords:
                continue

            lat1, lon1 = location_coords[sheet_name]

            # Find the nearest street address to this splice location using
            # brute-force comparison across all extracted addresses.
            best_addr = None
            best_dist = float("inf")
            for addr in addresses:
                dist = haversine(lat1, lon1, addr["Latitude"], addr["Longitude"])
                if dist < best_dist:
                    best_addr = addr["Address"]
                    best_dist = dist

            # Write the best address into cell B5 of the sheet.
            if best_addr:
                wb[sheet_name]["B5"] = best_addr
                updated_count += 1

        wb.save(FINAL_OUTPUT_PATH)
        print(f"Updated {updated_count} sheets.")
        print(f"Output saved to: {FINAL_OUTPUT_PATH}")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
