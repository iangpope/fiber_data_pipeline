#!/usr/bin/env python3
"""
0_run_pipeline_with_checkpoint.py -- Orchestrate the full Fiber Data Pipeline.

Runs steps 1 through 10 in sequence via pipeline_runner.run_pipeline(), pausing
after step 2 so the user can open and verify the Colored Connections Table before
committing to the full color-coding run.

Usage examples
--------------
  python3 0_run_pipeline_with_checkpoint.py               # full run with prompt
  python3 0_run_pipeline_with_checkpoint.py --yes         # skip prompt, auto-continue
  python3 0_run_pipeline_with_checkpoint.py --start 3     # resume from step 3
  python3 0_run_pipeline_with_checkpoint.py --start 5 --stop 7  # run steps 5-7 only

Exit codes
----------
  0 -- success (pipeline complete or stopped at checkpoint by user choice)
  1 -- a step failed
  2 -- bad arguments
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Checkpoint callback
# ---------------------------------------------------------------------------

CHECKPOINT_FILE = Path("output") / "Colored_Connections_Table.xlsx"


def ask_checkpoint(checkpoint_file: str, assume_yes: bool) -> bool:
    """
    Pause execution and prompt the user to verify the Colored Connections Table.

    The checkpoint is the critical review point: if direction colors are wrong
    at this stage, every downstream step will inherit the error.  The user should
    open the file, confirm each cable has the correct cardinal color, then type
    'y' to continue.

    Returns True to continue the pipeline, False to stop here.
    """
    print("\nCHECKPOINT")
    file_path = Path(checkpoint_file)
    if file_path.exists():
        print(f"Created: {checkpoint_file}")
    else:
        print(f"Expected file not found: {checkpoint_file}")
        print("(Step 2 should normally create it; check output above.)")

    print("\nPlease open the Colored Connections Table and verify direction colors.")
    print("Color guide: orange=North, green=East, brown=South, slate=West, red=MST, olive=OLT")
    print("When satisfied, type 'y' to continue, or anything else to stop.")

    if assume_yes:
        print("--yes flag set, continuing automatically.")
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
        help="Start from this step number (1-10). Useful to resume after a failure.",
    )
    parser.add_argument(
        "--stop",
        type=int,
        default=10,
        help="Stop after this step number (1-10).",
    )

    args = parser.parse_args()

    if not (1 <= args.start <= 10) or not (1 <= args.stop <= 10) or args.start > args.stop:
        print("Invalid --start/--stop range. Use 1-10 and ensure start <= stop.")
        return 2

    # Build the checkpoint callback with the --yes flag baked in.
    assume_yes = args.yes
    def checkpoint_callback(checkpoint_file: str) -> bool:
        return ask_checkpoint(checkpoint_file, assume_yes=assume_yes)

    # Import and run the pipeline.
    import pipeline_runner
    result = pipeline_runner.run_pipeline(
        data_dir="data",
        output_dir="output",
        start=args.start,
        stop=args.stop,
        checkpoint_callback=checkpoint_callback,
    )

    if result["status"] == "failed":
        print(f"\nPipeline failed: {result['error']}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
