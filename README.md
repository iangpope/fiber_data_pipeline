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

| File | Description |
|---|---|
| `*.kmz` | KMZ network export from the GIS design tool (exactly one) |
| `*.xlsx` | Raw cut sheet exported from the design tool (exactly one) |

The pipeline auto-detects both files by extension — no renaming required.

---

## Running the Pipeline

```bash
python3 0_run_pipeline_with_checkpoint.py
```

The runner executes all 8 steps in order and **pauses after step 2** so you can
open `output/Colored_Connections_Table.xlsx` and verify that every cable has been
assigned the correct directional color before the color-coding is applied to the
full workbook.

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
| 8 | `8_finalize.py` | Reordered with OTE | `output/Asbuilt_Workbook_post12.xlsx` ✅ |

The final deliverable is **`output/Asbuilt_Workbook_post12.xlsx`**.

---

## Color Key

| Color | Meaning |
|---|---|
| 🟠 Orange | North |
| 🟤 Brown | South |
| 🟢 Green | East |
| 🔵 Slate | West |
| 🔴 Red | MST tap |
| 🫒 Olive | OLT connection |
| 🟡 Yellow | Fusion splice / PORT boundary |
| 💙 Light blue | Splitter COMMON port |
| 🩷 Light pink | 1×32 splitter output |
| 🩷 Dark pink | 1×2 splitter output |
| 🍑 Peach | MUX port |
| 🐟 Salmon | DEMUX port |

---

## Shared Modules

| Module | Purpose |
|---|---|
| `config.py` | Central color palette, directory paths, and shared `PatternFill` cache |
| `naming_utils.py` | Location name parsing and classification (new `_FT_`/`_SE_` and legacy `MIC...` formats) |

All pipeline scripts import from these modules rather than defining their own
constants, ensuring consistent colors and behavior across every step.

---

## Project Structure

```
Fiber Data Pipeline/
├── data/                          # Input files (KMZ + cut sheet)
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
└── 8_finalize.py
```

---

## Branch Notes

This is the **`dev` branch**. It supports the new `RC73E_FT_001` / `RC73E_SE_001`
naming convention introduced for the RC73E project while retaining backward
compatibility with the legacy `MIC...` naming used in earlier projects.

Key improvements over `main`:
- Scripts renumbered 0–8 (no gaps) with a single orchestrating runner
- MST tap detection uses nearest-FT-per-SE logic instead of a fixed distance threshold
- Column detection is header-based rather than hardcoded, making scripts resilient
  to column reordering in future cut sheet exports
- All colors, paths, and shared logic centralized in `config.py` and `naming_utils.py`
- Professional commenting throughout all scripts
