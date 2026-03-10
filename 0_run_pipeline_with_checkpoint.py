#!/usr/bin/env python3
"""
0_run_pipeline_with_checkpoint.py -- Orchestrate the full Fiber Data Pipeline.

Runs steps 1 through 8 in sequence, pausing after step 2 so the user can
open and verify the Colored Connections Table before committing to the full
color-coding run. Each step is launched as a child process via subprocess so
that import-time code in scripts like 4_format_top_section.py does not
interfere with the runner's own namespace.

Usage examples:
  python3 0_run_pipeline_with_checkpoint.py               # full run with prompt
  python3 0_run_pipeline_with_checkpoint.py --yes         # skip prompt, auto-continue
  python3 0_run_pipeline_with_checkpoint.py --start 3     # resume from step 3
  python3 0_run_pipeline_with_checkpoint.py --start 5 --stop 7  # run steps 5-7 only

Exit codes:
  0 -- success (pipeline complete or stopped at checkpoint by user choice)
  1 -- a script failed (subprocess non-zero return code)
  2 -- bad arguments or missing script files
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Pipeline step order
# Each entry is a filename relative to this script's directory.
# Step numbers are 1-based (the index + 1 in this list).
# ---------------------------------------------------------------------------
SCRIPT_ORDER = [
    "1_coords_from_kmz.py",           # Step 1: Extract GPS coords from KMZ -> Connections_Table
    "2_compute_all_directions.py",     # Step 2: Compute cable bearings -> Colored_Connections_Table
    "3_colorize.py",                   # Step 3: Color cut sheet (main + splitter sections)
    "4_format_top_section.py",         # Step 4: Insert metadata + direction bars at top
    "5_change_connections_column.py",  # Step 5: Normalize CONNECTION column, shift/trim cols
    "6_assign_addresses.py",           # Step 6: Match splice locations to nearest street address
    "7_process_taps.py",               # Step 7: Reorder sheath blocks, shift port cols, label B3
    "8_finalize.py",                   # Step 8: Fill enclosure labels, insert rows, trim columns
]

# The manual checkpoint occurs after this script finishes (step 2).
# The user must open the output file and confirm all direction colors are correct
# before the remaining steps (which apply those colors throughout the workbook) run.
CHECKPOINT_AFTER = "2_compute_all_directions.py"
CHECKPOINT_FILE  = Path("output") / "Colored_Connections_Table.xlsx"


# ---------------------------------------------------------------------------
# Core execution helpers
# ---------------------------------------------------------------------------

def run_one(script_name: str) -> None:
    """
    Execute a single pipeline script as a child process.

    Using subprocess rather than importlib or exec() keeps each script's
    global-scope imports isolated from the runner and from each other,
    preventing side-effects between steps (e.g. module-level workbook loads
    in step 4 that would fail if the output from step 3 didn't exist yet).

    Raises RuntimeError if the script exits with a non-zero return code.
    """
    print(f"\nRunning {script_name} ...")
    proc = subprocess.run([sys.executable, script_name], text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"Script failed with code {proc.returncode}: {script_name}")


def ask_checkpoint(assume_yes: bool) -> bool:
    """
    Pause execution and prompt the user to verify the Colored Connections Table.

    The checkpoint is the critical review point in the pipeline. If direction
    colors are wrong at this stage, every downstream step will inherit the error.
    The user should open the file, confirm each cable has the correct cardinal
    color (orange=North, green=East, brown=South, slate=West, red=MST, olive=OLT),
    and then type 'y' to continue.

    Passing --yes on the command line skips the prompt and auto-continues (useful
    for automated testing or when the user has already verified the output).

    Returns True to continue, False to stop here.
    """
    print("\nCHECKPOINT")
    if CHECKPOINT_FILE.exists():
        print(f"Created: {CHECKPOINT_FILE}")
    else:
        print(f"Expected file not found yet: {CHECKPOINT_FILE}")
        print("(Step 2 should normally create it; check console output above.)")

    print("\nPlease open the Colored Connections Table and verify direction colors.")
    print("When you're happy, type 'y' to continue, or anything else to stop.")

    if assume_yes:
        print("--yes set, continuing automatically.")
        return True

    try:
        ans = input("Continue? [y/N]: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print("\nStopping.")
        return False

    return ans in {"y", "yes"}


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run the Fiber Data Pipeline with a manual checkpoint after step 2."
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Do not prompt at the checkpoint; continue automatically.",
    )
    parser.add_argument(
        "--start",
        type=int,
        default=1,
        help="Start from this step number (1-8). Useful to resume after a failure.",
    )
    parser.add_argument(
        "--stop",
        type=int,
        default=8,
        help="Stop after this step number (1-8).",
    )

    args = parser.parse_args()

    # Validate step range arguments.
    if not (1 <= args.start <= 8) or not (1 <= args.stop <= 8) or args.start > args.stop:
        print("Invalid --start/--stop range. Use 1-8 and ensure start <= stop.")
        return 2

    # Change to the directory containing this file so all relative paths
    # (script names, data/, output/) resolve correctly.
    os.chdir(Path(__file__).resolve().parent)

    # Verify all required script files exist before starting to avoid partial runs.
    missing = [s for s in SCRIPT_ORDER if not Path(s).exists()]
    if missing:
        print("Missing scripts:")
        for s in missing:
            print(f"  - {s}")
        return 2

    start_idx = args.start - 1
    stop_idx  = args.stop  - 1

    try:
        for i, script_name in enumerate(SCRIPT_ORDER[start_idx : stop_idx + 1], start=args.start):
            run_one(script_name)

            # After step 2, pause for manual review unless the user is stopping at step 2 anyway.
            if script_name == CHECKPOINT_AFTER and i != args.stop:
                if not ask_checkpoint(assume_yes=args.yes):
                    print("\nStopped at checkpoint (steps 3-8 were not run).")
                    return 0

        print("\nPipeline complete.")
        return 0

    except RuntimeError as e:
        print(f"\n{e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
