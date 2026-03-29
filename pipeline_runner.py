"""
pipeline_runner.py -- Callable pipeline module for the Fiber Data Pipeline.

Wraps all 10 pipeline steps into a single importable module so a web interface
or any other caller can invoke the full pipeline (or a subset of steps) without
shelling out via subprocess.

Public API
----------
    run_pipeline(data_dir, output_dir, start, stop,
                 checkpoint_callback, log_callback) -> dict

The checkpoint_callback is called after step 2 with the path to the Colored
Connections Table.  Return True from the callback to continue, False to stop.
If no callback is provided, the pipeline continues automatically.

The log_callback receives (step_number: int, line: str) for every line of
stdout produced by each step.  If omitted, output goes to sys.stdout as normal.

Return value
------------
{
    "status":      "complete" | "stopped_at_checkpoint" | "failed",
    "failed_step": int | None,
    "error":       str | None,
}
"""

from __future__ import annotations

import io
import os
import sys
import traceback
import contextlib
from pathlib import Path
from typing import Callable, Optional


# ---------------------------------------------------------------------------
# Step module imports
#
# Each step module now has a main(data_dir, output_dir) function.
# We import them lazily inside run_pipeline so that importing pipeline_runner
# itself has zero side effects.
# ---------------------------------------------------------------------------

STEP_MODULES = [
    "1_coords_from_kmz",
    "2_compute_all_directions",
    "3_colorize",
    "4_format_top_section",
    "5_change_connections_column",
    "6_assign_addresses",
    "7_process_taps",
    "8_finalize",
    "9_path_of_light",
    "10_generate_tap_report",
]

# Step number at which the manual checkpoint occurs (after this step finishes).
CHECKPOINT_AFTER_STEP = 2


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resolve_dirs(data_dir: str, output_dir: str) -> tuple[Path, Path]:
    """Return resolved absolute Paths for data and output directories."""
    base = Path(__file__).resolve().parent
    d = Path(data_dir)  if Path(data_dir).is_absolute()  else base / data_dir
    o = Path(output_dir) if Path(output_dir).is_absolute() else base / output_dir
    return d, o


def _import_step(module_name: str):
    """
    Import a pipeline step module by name and return it.

    Python module names cannot start with a digit, so step files like
    '1_coords_from_kmz.py' are imported via importlib with the filename
    used as the module spec source.
    """
    import importlib.util
    base = Path(__file__).resolve().parent
    spec = importlib.util.spec_from_file_location(
        module_name,
        str(base / f"{module_name}.py"),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _capture_and_relay(
    func,
    step_num: int,
    log_callback: Optional[Callable[[int, str], None]],
) -> None:
    """
    Run func(), capturing its stdout output and relaying each line to
    log_callback if provided, or printing it directly if not.
    """
    if log_callback is None:
        func()
        return

    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            func()
    finally:
        # Relay whatever was printed before the exception (or everything on success)
        output = buf.getvalue()
        for line in output.splitlines():
            log_callback(step_num, line)


# ---------------------------------------------------------------------------
# Main public API
# ---------------------------------------------------------------------------

def run_pipeline(
    data_dir:            str = "data",
    output_dir:          str = "output",
    start:               int = 1,
    stop:                int = 10,
    checkpoint_callback: Optional[Callable[[str], bool]] = None,
    log_callback:        Optional[Callable[[int, str], None]] = None,
) -> dict:
    """
    Run the Fiber Data Pipeline from step ``start`` to step ``stop`` inclusive.

    Parameters
    ----------
    data_dir : str
        Directory containing the input KMZ, cut sheet, and HAF files.
        Defaults to "data" relative to this file's location.
    output_dir : str
        Directory for all intermediate and final output workbooks.
        Defaults to "output" relative to this file's location.
    start : int
        First step to run (1–10, inclusive).
    stop : int
        Last step to run (1–10, inclusive).
    checkpoint_callback : callable(file_path: str) -> bool, optional
        Called after step 2 (direction assignment) with the path to the
        Colored Connections Table. Return True to continue, False to abort.
        If None, the pipeline continues automatically without pausing.
    log_callback : callable(step: int, line: str) -> None, optional
        Called with each line of stdout produced by each step. Useful for
        streaming progress to a web interface. If None, output goes directly
        to sys.stdout.

    Returns
    -------
    dict with keys:
        status       -- "complete", "stopped_at_checkpoint", or "failed"
        failed_step  -- step number that raised an exception, or None
        error        -- exception message if status is "failed", else None
    """
    if not (1 <= start <= 10) or not (1 <= stop <= 10) or start > stop:
        return {
            "status":      "failed",
            "failed_step": None,
            "error":       f"Invalid step range: start={start}, stop={stop}. Must be 1–10 with start <= stop.",
        }

    data_path, output_path = _resolve_dirs(data_dir, output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    # Change CWD to the pipeline directory so relative path references inside
    # step scripts (e.g. data/, output/) continue to resolve correctly.
    os.chdir(Path(__file__).resolve().parent)

    for step_num in range(start, stop + 1):
        module_name = STEP_MODULES[step_num - 1]

        if log_callback:
            log_callback(step_num, f"--- Step {step_num}: {module_name} ---")
        else:
            print(f"\nRunning step {step_num}: {module_name} ...")

        try:
            mod = _import_step(module_name)

            def _run(m=mod, d=str(data_path), o=str(output_path)):
                m.main(d, o)

            _capture_and_relay(_run, step_num, log_callback)

        except Exception as exc:
            tb   = traceback.format_exc()
            msg  = f"Step {step_num} ({module_name}) failed: {exc or type(exc).__name__}"
            if log_callback:
                log_callback(step_num, f"ERROR: {msg}")
                for line in tb.splitlines():
                    log_callback(step_num, line)
            else:
                print(f"\n{msg}\n{tb}", file=sys.stderr)
            return {
                "status":      "failed",
                "failed_step": step_num,
                "error":       f"{msg}\n{tb}",
            }

        # After step 2, invoke the checkpoint callback if we are not stopping here.
        if step_num == CHECKPOINT_AFTER_STEP and step_num < stop:
            checkpoint_file = str(output_path / "Colored_Connections_Table.xlsx")

            if checkpoint_callback is not None:
                should_continue = checkpoint_callback(checkpoint_file)
            else:
                should_continue = True   # no callback = auto-continue

            if not should_continue:
                if log_callback:
                    log_callback(step_num, "Stopped at checkpoint by user.")
                else:
                    print("\nStopped at checkpoint (steps 3-10 were not run).")
                return {
                    "status":      "stopped_at_checkpoint",
                    "failed_step": None,
                    "error":       None,
                }

    if log_callback:
        log_callback(stop, "Pipeline complete.")
    else:
        print("\nPipeline complete.")

    return {
        "status":      "complete",
        "failed_step": None,
        "error":       None,
    }
