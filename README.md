# Fiber Data Pipeline

Automated fiber optic data processing for FTTH projects. Takes a KMZ network map and a splice report (Excel cut sheet) as inputs, and produces a fully formatted, color-coded workbook for use as a field splicing guide.

---

## Overview

This pipeline processes raw outputs from network design tools and transforms them into a finalized as-built workbook. Each sheet in the output represents a splice location and contains:

- Color-coded sheath connections organized by cardinal direction (North/South/East/West)
- Highlighted optical splitter sections with port-type color coding
- Auto-populated metadata (enclosure type, tray count, GPS coordinates, nearest address)
- Reordered and cleaned tap records with PORT OTE labels

The pipeline runs as a series of numbered Python scripts (steps 1–12), orchestrated by a single runner with a built-in review checkpoint.

---

## Inputs

Place the following files in the `data/` directory before running:

| File | Description |
|---|---|
| `*.kmz` | KMZ network map exported from your GIS/design tool (one file) |
| `*.xlsx` | Splice report / cut sheet exported from your splice management system (one file) |

The scripts auto-detect both files by extension — no renaming required.

---

## Output

All output files are written to the `output/` directory. The final deliverable is:

```
output/Asbuilt_Workbook_post12.xlsx
```

Intermediate files from each step are also preserved in `output/` for debugging.

---

## Pipeline Steps

| Step | Script | Description |
|---|---|---|
| 1 | `1_coords_from_kmz.py` | Parses KMZ placemarks to extract GPS coordinates for each splice location; builds `Connections_Table.xlsx` |
| 2 | `2_compute_all_directions.py` | Computes cable bearing from GPS geometry; assigns directional colors (N/S/E/W/OLT/MST) to each connection; outputs `Colored_Connections_Table.xlsx` |
| — | **CHECKPOINT** | Pause here to manually verify direction colors in the connections table before continuing |
| 3 | `3_color_cut_sheet.py` | Applies sheath colors to the cut sheet; color-codes splitter output ports (COMMON/OUT/MUX/DEMUX) on splitter sheets |
| 4 | `4_update_and_highlight_splitter_section.py` | Colors rows within the OPTICAL SPLITTERS section by port type; highlights first row of each device group in yellow |
| 5 | `5_format_top_section.py` | Inserts metadata table (splice ID, enclosure, coordinates) and directional color bars at the top of each sheet |
| 6 | `6_change_connections_column.py` | Normalizes the CONNECTION column (replaces CONTINUOUS/FUSION with `<--->`, adjusts formatting); shifts and trims columns in splitter sections |
| 7 | `7_assign_addresses.py` | Matches each splice location to the nearest street address from KMZ folder placemarks using haversine distance |
| 8 | `8_reorder_sheaths_detect_M_or_N.py` | On tap sheets, reorders sheath blocks so the block containing PORT rows appears first |
| 9 | `9_shift_ports_preserve_all_ports.py` | Shifts PORT NAME / PORT WAVELENGTH / DEVICE NAME columns leftward into a consistent position; trims excess columns |
| 10 | `10_label_b3_final_shifted_fixed_cols.py` | Detects port count on tap sheets and writes the enclosure label (e.g. `2 PORT OTE`, `4 PORT OTE`) to cell B3 |
| 11 | `11_final.py` | Final cleanup: fills missing enclosure/tray values based on naming convention; inserts MST and DEMUX lines where applicable |
| 12 | `12_cleanup_splitter_and_trim.py` | On splitter sheets, shifts per-splitter metadata leftward and trims excess columns; clears SHEATH UUID from the splitter sub-table |

---

## Color Reference

### Directional (sheath connections)

| Color | Meaning |
|---|---|
| 🟠 Orange (`FFA500`) | North |
| 🟤 Brown (`8B4513`) | South |
| 🟢 Green (`008000`) | East |
| 🔵 Slate (`708090`) | West |
| 🟥 Red (`FF0000`) | MST (main supply trunk) |
| 🌿 Olive Green (`C5D9B5`) | OLT connection |

### Splitter Port Types

| Color | Meaning |
|---|---|
| 🔵 Light Blue (`ADD8E6`) | COMMON port |
| 🩷 Light Pink (`FFB6C1`) | 1×32 splitter output |
| 💗 Dark Pink (`DB7093`) | 1×2 splitter output |
| 🍑 Peach (`FFDAB9`) | MUX channel port |
| 🍊 Salmon (`FFA07A`) | DEMUX channel port |
| 🟡 Yellow (`FFFF00`) | FUSION splice / first row of device group |

---

## Running the Pipeline

### Run all steps (with checkpoint)

```bash
cd "Fiber Data Pipeline"
python 0_run_pipeline_with_checkpoint.py
```

The pipeline pauses after step 2 and asks you to verify the direction color assignments in `output/Colored_Connections_Table.xlsx`. Type `y` to continue.

### Skip the checkpoint prompt

```bash
python 0_run_pipeline_with_checkpoint.py --yes
```

### Run from a specific step (e.g., resume from step 5)

```bash
python 0_run_pipeline_with_checkpoint.py --start 5
```

### Run only a range of steps

```bash
python 0_run_pipeline_with_checkpoint.py --start 3 --stop 6
```

---

## Requirements

Install dependencies with:

```bash
pip install openpyxl pandas geopy
```

Python 3.9+ recommended.

---

## Supported Sheet Naming Conventions

The pipeline handles both legacy and current location naming formats:

| Format | Example | Type |
|---|---|---|
| `{OLT}E` | `CD60E` | OLT node |
| `MIC{OLT}S{NNN}` | `MICCD60S001` | Splitter (legacy) |
| `MIC{OLT}D{NNN}` | `MICMS36D002` | Distribution (legacy) |
| `{OLT}E_FT_{NNN}` | `CD60E_FT_032` | Tap enclosure |
| `{OLT}E_SE_{NNN}` | `CD60E_SE_001` | Tap enclosure |

---

## Project Structure

```
Fiber Data Pipeline/
├── data/                        # Input files (gitignored)
│   ├── *.kmz                    # KMZ network map
│   └── *.xlsx                   # Source cut sheet / splice report
├── output/                      # Generated files (gitignored)
├── 0_run_pipeline_with_checkpoint.py
├── 1_coords_from_kmz.py
├── 2_compute_all_directions.py
├── 3_color_cut_sheet.py
├── 4_update_and_highlight_splitter_section.py
├── 5_format_top_section.py
├── 6_change_connections_column.py
├── 7_assign_addresses.py
├── 8_reorder_sheaths_detect_M_or_N.py
├── 9_shift_ports_preserve_all_ports.py
├── 10_label_b3_final_shifted_fixed_cols.py
├── 11_final.py
├── 12_cleanup_splitter_and_trim.py
└── naming_utils.py              # Shared location name parsing utilities
```

---

## Branch Strategy

| Branch | Purpose |
|---|---|
| `main` | Stable, production-tested pipeline |
| `dev` | Active development and refactoring |

All refactoring work is done on `dev` and merged to `main` only when verified.
