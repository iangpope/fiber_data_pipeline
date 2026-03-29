# Fiber Data Pipeline — Dev Branch

> **Automated FTTH splice workbook generator.**
> Converts KMZ network geometry and a raw cut sheet into a fully color-coded,
> address-labeled, field-ready splice guide.

---

## Overview

This pipeline processes GIS exports from a fiber network design tool and produces
a finalized Excel workbook used by field technicians as a splicing guide.
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
| `flask` | Web UI server |

Install dependencies:

```bash
pip install openpyxl pandas geopy flask
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

### Web UI (recommended)

```bash
cd tools/pipeline_ui
./start.sh          # or: python3 run.py
```

Open `http://localhost:5000` in a browser. Upload the KMZ and cut sheet (and
optionally a HAF report), then click **Run Pipeline**. The UI streams live
progress, redirects to the checkpoint review page after step 2, and provides
download links on completion.

#### Checkpoint Review Page

After steps 1–2 complete, the pipeline pauses at a checkpoint with:

- **Interactive Leaflet map** showing the fiber cable network colored by
  direction. Only fiber cables are shown (support/infrastructure conduits are
  hidden for clarity). Click any splice location node to highlight the cables
  attached to it and display their sheath labels.
- **Direction override** — click a cable color chip to reassign its compass
  direction (North / South / East / West / OLT / MST). Changes are written back
  to the Colored Connections Table immediately.
- **Cable resize** — change a cable's CT count (12 / 24 / 48 / 96 / 144). The
  tool renames all occurrences in both the Connections Table and the cut sheet,
  and also adjusts the fiber row count in the cut sheet:
  - New rows continue the TIA-598 buffer/fiber color sequence (BL→OR→GR→BR→SL→WH→RD→BK→YE→VI→PI→AQ)
  - Rows are marked connected (`<- FUSION ->`) up to the fiber count the
    right-side cable supports; remaining rows are marked `X` with blank right-side columns
  - When both cables at a splice are resized, the cross-sheet reconciliation
    pass automatically heals previously-X rows in sibling cable blocks

### Command Line (headless)

```bash
python3 0_run_pipeline_with_checkpoint.py
```

The runner executes steps 1–8 in order and **pauses after step 2** so you can
open `output/Colored_Connections_Table.xlsx` and verify cable direction colors
before applying them to the full workbook.

#### Options

```
--yes           Skip the checkpoint prompt and continue automatically
--start N       Resume from step N (1–8)
--stop N        Stop after step N (1–8)
```

---

## Pipeline Steps

| # | Script | Input | Output |
|---|---|---|---|
| 1 | `1_coords_from_kmz.py` | KMZ, cut sheet | `data/Connections_Table.xlsx` |
| 2 | `2_compute_all_directions.py` | KMZ, Connections Table | `output/Colored_Connections_Table.xlsx` |
| **—** | **Checkpoint** | Verify / adjust direction colors | |
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
generates the field Tap Report from the HAF address data and asbuilt workbook.

---

## Direction Algorithm

Step 2 assigns a compass direction (North / East / South / West) to each cable
at each splice location using:

1. **KMZ geometry walk** — walks 25 m along each cable's actual routed path from
   the splice node, rather than using the straight endpoint-to-endpoint bearing.
   This gives accurate directions even for cables that leave a node at an angle
   before routing in a different compass direction.

2. **Optimal assignment** — when multiple cables leave the same node, an
   exhaustive search over all direction permutations minimizes total angular
   deviation. This prevents the suboptimal assignments that arise from greedy
   first-come-first-served direction picking at complex junctions.

Step 2 also emits a confidence log listing cables that are short (< 25 m, less
reliable bearing) or not found in the KMZ geometry.

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
| Light pink | 1×32 splitter output |
| Dark pink | 1×2 splitter output |
| Peach | MUX port |
| Salmon | DEMUX port |

---

## Shared Modules

| Module | Purpose |
|---|---|
| `config.py` | Central color palette, `PatternFill` cache, directory paths, TIA-598 fiber/buffer color sequence, and connection column value constants |
| `naming_utils.py` | Location name parsing and classification (`_FT_`/`_SE_` and legacy `MIC...` formats); worksheet column/row detection utilities |

---

## Tools

### Pipeline Web UI — `tools/pipeline_ui/`

Flask web application that wraps the full pipeline with:
- Drag-and-drop file upload for KMZ, cut sheet, and optional HAF report
- Real-time progress via Server-Sent Events (spinner + step label during normal
  operation; terminal log revealed only on failure)
- Interactive checkpoint review map (Leaflet.js) with direction overrides and
  cable resize
- Download links for all output files on completion

Start: `cd tools/pipeline_ui && ./start.sh` (default port 5000)

### Cable Resize Tool — `tools/resize_cables/`

Standalone Flask tool for resizing cables in an existing cut sheet outside of a
pipeline run. Provides a simple web form to select a cable and target CT count.

Start: `cd tools/resize_cables && ./start.sh` (default port 5001)

---

## Project Structure

```
Fiber Data Pipeline/
├── data/                          # Input files (KMZ, cut sheet, HAF, template)
├── output/                        # All intermediate and final outputs
├── config.py                      # Shared constants, color fills, TIA-598 sequence
├── naming_utils.py                # Location naming and classification
├── pipeline_runner.py             # Orchestrates steps 1–10 with callbacks
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
├── 10_generate_tap_report.py
└── tools/
    ├── pipeline_ui/               # Web UI (Flask + Leaflet checkpoint map)
    │   ├── app.py
    │   ├── map_builder.py
    │   ├── static/
    │   └── templates/
    └── resize_cables/             # Standalone cable resize tool
        ├── app.py
        ├── resize_logic.py
        ├── static/
        └── templates/
```

---

## Branch Notes

This is the **`dev` branch**. It supports the `RC73E_FT_001` / `RC73E_SE_001`
naming convention introduced for the RC73E project while retaining backward
compatibility with the legacy `MIC...` naming used in earlier projects.

Key improvements over `main`:
- **Pipeline Web UI** with SSE log streaming, interactive checkpoint map, inline
  direction overrides, and cable CT resizing with automatic row adjustment
- **Improved direction algorithm** — KMZ geometry walk (25 m) + exhaustive
  optimal assignment replaces endpoint bearing + greedy deduplication
- **Cable resize row adjustment** — extending a cable's CT count adds rows with
  correct TIA-598 buffer/fiber sequence and proper connection or X markers;
  cross-sheet reconciliation heals sibling cable blocks when both cables at a
  splice are resized
- Scripts renumbered 0–10 with a single orchestrating runner (`pipeline_runner.py`)
- MST tap detection uses nearest-FT-per-SE logic instead of a fixed distance threshold
- Column detection is header-based throughout — resilient to column additions or
  reordering in future Magellan exports
- Connection column value strings centralized in `config.py` as named constants
- All colors, paths, and shared logic centralized in `config.py` and `naming_utils.py`
