# AI Context — Fiber Data Pipeline

Read this file before making any changes to the project.

---

## What This Project Does

Automated FTTH splice workbook generator. Takes a KMZ network geometry export
and a raw cut sheet (Excel) from the GIS design tool and produces a finalized,
color-coded Excel workbook used by field technicians as a splicing model and guide.

---

## Naming Conventions

Two location naming systems are in use in the same project:

| Format | Example | Meaning |
|---|---|---|
| New | `RC73E_FT_001` | FT = tap/OTE enclosure |
| New | `RC73E_SE_006` | SE = splice enclosure (splitter or distribution) |
| Legacy | `MICMS02S007` | Prefix + S/D type hint + number |
| OLT token | `RC73E` | Bare site ID, no suffix — the OLT rack sheet |

All naming parsing is centralized in `naming_utils.py`. Do not duplicate it.

---

## Network Topology

```
OLT rack (RC73E)
  └─> backhaul cable
        └─> SE enclosure  (OPTICAL SPLITTERS section)
              ├─> COMMON port  = upstream backhaul cable
              └─> Out-N ports  = downstream cables to taps or more SEs
                    └─> FT tap enclosure  (SHEATHS section, PORT rows)
```

Key rules:
- An SE enclosure **always** has an `OPTICAL SPLITTERS` sub-table if it contains a splitter.
- An SE can **also** be a pass-through/ring node whose SHEATHS section carries cables
  that don't interact with its local splitter at all (e.g. SE_011 is pass-through for the
  FT_059–FT_068 branch even though it hosts its own local splitter).
- Splitters can be **cascaded**: Out-N of one device (e.g. 1×2) feeds directly into
  the COMMON of another (e.g. 1×32) without an intermediate cable sheath. The
  OPTICAL SPLITTERS rows show this via device names in the right-side columns (E/F)
  instead of a cable name in column K.

---

## Workbook Column Layout

### SHEATHS section header row

| Col A | Col B | Col C | Col D | Col E | Col F | Col G | ... | Col J | Col K | Col L |
|---|---|---|---|---|---|---|---|---|---|---|
| SHEATH UUID | SHEATH NAME | START ENC | END ENC | BUFFER | FIBER | CONNECTION | ... | (right) START ENC | (right) END ENC | (right) SHEATH NAME |

- `<--->` in CONNECTION = fused/active splice
- `X` in CONNECTION = unused fiber

On **tap (FT) sheets**, after step 7 (`7_process_taps.py`), **column J** holds the
PORT label (e.g. `PORT1`, `PORT2`) for PORT rows.

### OPTICAL SPLITTERS sub-table header row

| Col A | Col B | Col C | Col D | Col E | Col F | Col G | Col H | Col I | Col J | Col K |
|---|---|---|---|---|---|---|---|---|---|---|
| DEVICE UUID | DEVICE NAME | PORT NAME (left) | CONNECTION | PORT NAME (right) | DEVICE NAME (right) | FIBER | BUFFER | END ENC | START ENC | SHEATH NAME |

- `Common` in PORT NAME = upstream input port (cable = backhaul toward OLT)
- `Out-N` in PORT NAME = downstream output port (cable = toward tap)
- If SHEATH NAME (col K) is **empty** on an Out-N or Common row, columns E/F hold the 
  device name of the other splitter it connects to (cascaded configuration).

---

## Shared Modules

| File | Purpose |
|---|---|
| `config.py` | Color palette, directory paths, `PatternFill` cache |
| `naming_utils.py` | All location name parsing, classification, and worksheet column-finding utilities |

**Always use `naming_utils.find_col_by_header()` and `find_header_row()` instead of
hardcoded column numbers.** The column layout can shift between projects.

---

## Pipeline Steps

| # | Script | What it does |
|---|---|---|
| 0 | `0_run_pipeline_with_checkpoint.py` | Orchestrator; pauses after step 2 for manual review |
| 1 | `1_coords_from_kmz.py` | Extract GPS coords + cable connections from KMZ |
| 2 | `2_compute_all_directions.py` | Calculate cable bearings; assign directional colors |
| 3 | `3_colorize.py` | Apply directional + splitter colors to cut sheet |
| 4 | `4_format_top_section.py` | Insert metadata block + direction bar legend |
| 5 | `5_change_connections_column.py` | Normalize CONNECTION column; trim splitter cols |
| 6 | `6_assign_addresses.py` | Match splice locations to nearest street address |
| 7 | `7_process_taps.py` | Reorder sheath blocks; shift PORT cols to J/K/L; label B3 |
| 8 | `8_finalize.py` | Fill enclosure labels; insert DEMUX/MST rows; trim columns |
| 9 | `9_path_of_light.py` | PON continuity tracer — traces every tap PORT to OLT |

Final output: `output/Asbuilt_Workbook_post12.xlsx`
Trace report: `output/Path_of_Light_Confirmation.xlsx`

---

## Path of Light Tracer Logic

`9_path_of_light.py` traces backward from each tap PORT to the OLT:

1. Read the cable name on the PORT row (SHEATHS section of tap sheet, col B)
2. Navigate to the SE sheet named in the END ENCLOSURE column of that row
3. At the SE: look up the cable in the OPTICAL SPLITTERS output map
   - If found → follow to the `Common` port cable → that is the next upstream cable
   - If **not** found → the cable is a SHEATHS pass-through at this SE (the SE is a
     ring/distribution node that also hosts a local splitter for other taps).
     Follow via `_trace_through_sheaths()` instead.
4. Repeat until node name matches the bare OLT token (e.g. `RC73E`)

Common failure modes the tracer catches:
- **Missing splitter output assignments** — cable exists in SHEATHS but no Out-N row
  references it in the OPTICAL SPLITTERS table (design omission)
- **Dead-end cable** — cable has no far-end node listed in the design data
- **Missing sheet** — a node referenced in the workbook has no corresponding sheet

---

## Color Key

| Color (hex) | Meaning |
|---|---|
| Orange | North direction |
| Brown | South direction |
| Green | East direction |
| Slate blue | West direction |
| Red | MST tap enclosure |
| Olive green | OLT connection |
| Yellow | Fusion splice / PORT boundary marker |
| Light blue | Splitter COMMON port |
| Light pink | 1×32 splitter output |
| Dark pink | 1×2 splitter output |
| Peach | MUX port |
| Salmon | DEMUX port |

All hex values are defined in `config.py` under the `COLOR` dict.

---

## Git Repository

**Remote:** `https://github.com/iangpope/fiber_data_pipeline.git`

```bash
git clone https://github.com/iangpope/fiber_data_pipeline.git
cd fiber_data_pipeline
git checkout dev   # active development branch
```

---

## Branch Strategy

| Branch | Purpose |
|---|---|
| `main` | Stable release — legacy naming convention only (MIC... format) |
| `dev` | Active development — supports both new `_FT_`/`_SE_` and legacy naming |

**Always work on `dev`.** Only merge to `main` when a full project has been
completed and tested end-to-end.

The pipeline is **project-agnostic** — it works with any OLT site. The OLT token
(e.g. `RC73E`, `MS90E`) is detected automatically from the sheet names and cable
naming in the input files. RC73E is the most recently processed project but the
code makes no assumptions about which site is being processed.

### Current state of `dev` (as of March 2026)

| Commit | Description |
|---|---|
| `218d5a7` | Add AGENTS.md |
| `989612b` | Fix tracer: SE pass-through at SE nodes that also host splitters |
| `beaef19` | Add `9_path_of_light.py` — PON continuity tracer |
| `8a08a90` | Add professional README |
| `d734a64` | Fix metadata header font (Calibri) |
| `4b2a2b5` | Professional commenting across all scripts |
| `5623781` | Renumber scripts 0–8; fix classify_sheet; clean terminal output |
| `c3c0fa7` | Remove superseded scripts (old 3, 4, 8, 9, 10, 11, 12) |

### Typical workflow

```bash
# Start a session
git checkout dev
git pull origin dev

# After making changes
git add <files>
git commit -m "Short description of what changed"
git push origin dev
```

