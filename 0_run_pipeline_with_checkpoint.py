#!/usr/bin/env python3
"""Run the AltColorFiberSheets pipeline with a manual checkpoint.

Runs steps 1-12 in order, but pauses after step 2 (which produces
output/Colored_Connections_Table.xlsx) so you can verify direction colors.

Examples:
  python3 0_run_pipeline_with_checkpoint.py
  python3 0_run_pipeline_with_checkpoint.py --yes
  python3 0_run_pipeline_with_checkpoint.py --start 3
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


SCRIPT_ORDER = [
    "1_coords_from_kmz.py",
    "2_compute_all_directions.py",
    "3_color_cut_sheet.py",
    "4_update_and_highlight_splitter_section.py",
    "5_format_top_section.py",
    "6_change_connections_column.py",
    "7_assign_addresses.py",
    "8_reorder_sheaths_detect_M_or_N.py",
    "9_shift_ports_preserve_all_ports.py",
    "10_label_b3_final_shifted_fixed_cols.py",
    "11_final.py",
    "12_cleanup_splitter_and_trim.py",
]

CHECKPOINT_AFTER = "2_compute_all_directions.py"  # step 2 generates the colored connections table
CHECKPOINT_FILE = Path("output") / "Colored_Connections_Table.xlsx"


def run_one(script_name: str) -> None:
    """Run a single script as a child process, raising on failure."""
    print(f"\n▶ Running {script_name} ...")
    proc = subprocess.run([sys.executable, script_name], text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"Script failed with code {proc.returncode}: {script_name}")


def ask_checkpoint(assume_yes: bool) -> bool:
    """Return True to continue, False to stop."""
    print("\n⏸  CHECKPOINT")
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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run AltColorFiberSheets pipeline with a checkpoint after step 2."
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
        help="Start from this step number (1-12). Useful to resume.",
    )
    parser.add_argument(
        "--stop",
        type=int,
        default=12,
        help="Stop after this step number (1-12).",
    )

    args = parser.parse_args()

    if not (1 <= args.start <= 12) or not (1 <= args.stop <= 12) or args.start > args.stop:
        print("Invalid --start/--stop range. Use 1-12 and ensure start <= stop.")
        return 2

    # Ensure we run from the folder containing this file
    os.chdir(Path(__file__).resolve().parent)

    # Quick sanity check: ensure scripts exist
    missing = [s for s in SCRIPT_ORDER if not Path(s).exists()]
    if missing:
        print("Missing scripts:")
        for s in missing:
            print(f"  - {s}")
        return 2

    start_idx = args.start - 1
    stop_idx = args.stop - 1

    try:
        for i, script_name in enumerate(SCRIPT_ORDER[start_idx : stop_idx + 1], start=args.start):
            run_one(script_name)

            if script_name == CHECKPOINT_AFTER and i != args.stop:
                # Pause after step 2 unless user is stopping anyway
                if not ask_checkpoint(assume_yes=args.yes):
                    print("\n✅ Stopped at checkpoint (nothing after step 2 was run).")
                    return 0

        print("\n✅ Pipeline complete.")
        return 0

    except RuntimeError as e:
        print(f"\n❌ {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
