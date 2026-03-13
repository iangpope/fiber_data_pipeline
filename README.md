# Fiber Data Pipeline — Dev Branch

> **Automated FTTH splice workbook generator.**
> Converts KMZ network geometry and a raw cut sheet into a fully color-coded,
> address-labeled, field-ready splice guide.

---

## Overview

This pipeline processes GIS exports from a fiber network design tool and produces
a finalized Excel workbook used by field technicians as a splicing model and guide.
Each sheet in the output corresponds to one splice enclosure or tap location and
contains:

- Color-coded sheath rows by compass direction (North / East / South / West)
- Red highlighting for MST (Main Sheath Terminal) tap enclosures
- Olive green for OLT (Optical Line Terminal) connections
- Port reordering and OTE size labeling for tap sheets
- Optical splitter port coloring (COMMON, 1×32, 1×2, MUX, DEMUX)
- Nearest street address auto-filled into each sheet
- Enclosure type and tray count labels (COMMSCOPE FOSC 450-B/D, PORT OTE sizes)

---

## Requirements

| Dependency | Purpose |
|---|---|
| Python 3.11+ | Runtime |
| `openpyxl` | Excel read/write |
| `pandas` | Tabular data processing |
| `geopy` | Geodesic distance calculation |

Install dependencies:

```bash
pip install openpyxl pandas geopy
```

---

## Input Files

Place the following files in the `data/` folder before running:

| File pattern | Description |
|---|---|
| `*.kmz` | KMZ network export from the GIS design tool (exactly one) |
| `*SPLICE*REPORTS*FIBER*.xlsx` | Raw cut sheet (Splice Reports) exported from Magellan (exactly one) |
| `*HAF*.xlsx` | HAF address report (required for step 10 only) |
| `Tap_Report_Template.xlsx` | Tap Report template (required for step 10 only) |

The pipeline auto-detects input files by name pattern — no renaming required.

---

## Running the Pipeline

```bash
python3 0_run_pipeline_with_checkpoint.py
```

The runner executes steps 1–8 in order and **pauses after step 2** so you can
open `output/Colored_Connections_Table.xlsx` and verify that every cable has been
assigned the correct directional color before the color-coding is applied to the
full workbook.

Step 2 also prints a **debug/confidence log** to the terminal listing:
- Cables not found in the KMZ (no geometry, no color assigned)
- Cables shorter than 5 m (bearing taken from far endpoint, less reliable)
- Cables whose bearing falls in a diagonal zone (direction assignment less certain)

Review these before approving the checkpoint.

Steps 9 and 10 are run independently after the main pipeline completes.

### Options

```
--yes           Skip the checkpoint prompt and continue automatically
--start N       Resume from step N (1–8), e.g. after fixing an input file
--stop N        Stop after step N (1–8), e.g. to inspect intermediate output
```

Examples:

```bash
python3 0_run_pipeline_with_checkpoint.py --yes           # fully automated
python3 0_run_pipeline_with_checkpoint.py --start 3       # re-run from step 3
python3 0_run_pipeline_with_checkpoint.py --start 5 --stop 7
```

---

## Pipeline Steps

| # | Script | Input | Output |
|---|---|---|---|
| 1 | `1_coords_from_kmz.py` | KMZ, cut sheet | `data/Connections_Table.xlsx` |
| 2 | `2_compute_all_directions.py` | KMZ, Connections Table | `output/Colored_Connections_Table.xlsx` |
| **—** | **Manual checkpoint** | Verify direction colors | |
| 3 | `3_colorize.py` | Cut sheet, Colored Connections | `output/Colorized_Cut_Sheet_Final_v7_highlighted.xlsx` |
| 4 | `4_format_top_section.py` | Colorized cut sheet | `output/Combined_Formatted_Output.xlsx` |
| 5 | `5_change_connections_column.py` | Combined formatted output | `output/Combined_Formatted_Output_processed.xlsx` |
| 6 | `6_assign_addresses.py` | KMZ, processed output | `output/Combined_Formatted_Output_with_Addresses.xlsx` |
| 7 | `7_process_taps.py` | Output with addresses | `output/Combined_Reordered_With_OTE.xlsx` |
| 8 | `8_finalize.py` | Reordered with OTE | `output/Asbuilt_Workbook_post12.xlsx` |
| 9 | `9_path_of_light.py` | Asbuilt workbook | `output/Path_of_Light_Confirmation.xlsx` |
| 10 | `10_generate_tap_report.py` | HAF report, asbuilt workbook | `output/{OLT Name} Tap Report.xlsx` |

Steps 1–8 produce the finished splice workbook. Step 9 verifies PON continuity
by tracing every tap PORT back to the OLT and reporting any broken paths. Step 10
generates the field Tap Report by combining the HAF address data with the burn
summary extracted from the asbuilt workbook.

---

## Color Key

| Color | Meaning |
|---|---|
| Orange | North |
| Brown | South |
| Green | East |
| Slate blue | West |
| Red | MST tap |
| Olive green | OLT connection |
| Yellow | Fusion splice / PORT boundary |
| Light blue | Splitter COMMON port |
| Light pink | 1x32 splitter output |
| Dark pink | 1x2 splitter output |
| Peach | MUX port |
| Salmon | DEMUX port |

---

## Shared Modules

| Module | Purpose |
|---|---|
| `config.py` | Central color palette, `PatternFill` cache, directory paths, and connection column value constants (`CONN_RAW_FUSION`, `CONN_FUSED`, etc.) |
| `naming_utils.py` | Location name parsing and classification (new `_FT_`/`_SE_` and legacy `MIC...` formats); worksheet column/row detection utilities |

All pipeline scripts import from these modules rather than defining their own
constants, ensuring consistent colors and behavior across every step.

---

## Project Structure

```
Fiber Data Pipeline/
├── data/                          # Input files (KMZ, cut sheet, HAF, template)
├── output/                        # All intermediate and final outputs
├── config.py                      # Shared constants and color fills
├── naming_utils.py                # Location naming and classification
├── 0_run_pipeline_with_checkpoint.py
├── 1_coords_from_kmz.py
├── 2_compute_all_directions.py
├── 3_colorize.py
├── 4_format_top_section.py
├── 5_change_connections_column.py
├── 6_assign_addresses.py
├── 7_process_taps.py
├── 8_finalize.py
├── 9_path_of_light.py
└── 10_generate_tap_report.py
```

---

## Branch Notes

This is the **`dev` branch**. It supports the new `RC73E_FT_001` / `RC73E_SE_001`
naming convention introduced for the RC73E project while retaining backward
compatibility with the legacy `MIC...` naming used in earlier projects.

Key improvements over `main`:
- Scripts renumbered 0–8 (no gaps) with a single orchestrating runner
- MST tap detection uses nearest-FT-per-SE logic instead of a fixed distance threshold
- Cable orientation uses geodesic nearest-endpoint comparison instead of exact float
  equality, correctly handling floating point differences between KMZ and table coordinates
- Column detection is header-based throughout steps 3, 5, 7, and 8 — resilient to
  column additions or reordering in future Magellan exports
- Metadata row detection (enclosure label, tray count) uses column-A content scan
  instead of fixed row numbers — resilient to changes in the metadata block height
- Connection column value strings centralized in `config.py` as named constants
- Step 2 emits a debug/confidence log flagging cables to scrutinize before checkpoint approval
- All colors, paths, and shared logic centralized in `config.py` and `naming_utils.py`
- Professional commenting throughout all scripts
