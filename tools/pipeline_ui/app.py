"""
app.py -- Pipeline Web UI Flask application.

Runs the full fiber data pipeline from the browser with:
  - Dual file upload (KMZ + cut sheet; optional HAF + tap template)
  - Real-time log streaming via Server-Sent Events (SSE)
  - Inline checkpoint review page with Leaflet map overlay (after step 2)
  - Download links for all output files
"""

from __future__ import annotations

import io
import json
import os
import queue
import shutil
import sys
import tempfile
import threading
import uuid
from pathlib import Path

import re
import openpyxl
from openpyxl.styles import PatternFill
from flask import (
    Flask, Response, redirect, render_template,
    request, send_file, session, url_for, flash, jsonify
)

# ---------------------------------------------------------------------------
# Path setup — pipeline root must be on sys.path for pipeline_runner import
# ---------------------------------------------------------------------------
_HERE = Path(__file__).resolve().parent          # tools/pipeline_ui/
_ROOT = _HERE.parent.parent                      # Fiber Data Pipeline/
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import pipeline_runner
from map_builder import build_geojson


def _rename_cable_size(cable: str, new_size: str) -> str:
    """Return the cable name with its CT suffix replaced by new_size."""
    m = re.search(r'_(\d{2,3}CT)$', cable, re.IGNORECASE)
    if m:
        return cable[:m.start()] + '_' + new_size.upper()
    m = re.match(r'^(\d{2,3}CT)(\s)', cable, re.IGNORECASE)
    if m:
        return new_size.upper() + cable[m.end(1):]
    return cable



# ---------------------------------------------------------------------------
# In-memory job registry
# Each job: { queue, thread, checkpoint_event, checkpoint_continue,
#             job_dir, status, outputs }
# ---------------------------------------------------------------------------
_JOBS: dict[str, dict] = {}


def _make_job_id() -> str:
    return uuid.uuid4().hex[:12]


# ---------------------------------------------------------------------------
# Pipeline background worker
# ---------------------------------------------------------------------------

def _run_job(job_id: str, data_dir: str, output_dir: str) -> None:
    job = _JOBS[job_id]
    q   = job["queue"]

    def log_cb(step: int, line: str) -> None:
        q.put({"type": "log", "step": step, "line": line})

    def checkpoint_cb(file_path: str) -> bool:
        """Pause the pipeline thread and wait for the browser to respond."""
        q.put({"type": "checkpoint", "path": file_path})
        job["checkpoint_event"].wait()          # block until user clicks Continue/Stop
        return job["checkpoint_continue"]       # True = continue, False = abort

    try:
        result = pipeline_runner.run_pipeline(
            data_dir=data_dir,
            output_dir=output_dir,
            start=1,
            stop=10,
            checkpoint_callback=checkpoint_cb,
            log_callback=log_cb,
        )
        job["status"] = result["status"]
        # Collect output files
        outputs = {}
        for fname in ["Asbuilt_Workbook_post12.xlsx", "Tap_Report.xlsx",
                      "Path_of_Light_Confirmation.xlsx"]:
            fpath = Path(output_dir) / fname
            if fpath.exists():
                outputs[fname] = str(fpath)
        job["outputs"] = outputs
        q.put({"type": "done", "status": result["status"],
               "error": result.get("error"), "outputs": list(outputs.keys())})
    except Exception as exc:
        job["status"] = "failed"
        q.put({"type": "done", "status": "failed", "error": str(exc), "outputs": []})


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.urandom(32)
    app.config["MAX_CONTENT_LENGTH"] = 128 * 1024 * 1024  # 128 MB

    # -----------------------------------------------------------------------
    # Landing / upload page
    # -----------------------------------------------------------------------
    @app.route("/", methods=["GET"])
    def index():
        return render_template("index.html")

    # -----------------------------------------------------------------------
    # Start a pipeline run
    # -----------------------------------------------------------------------
    @app.route("/run", methods=["POST"])
    def start_run():
        kmz_file   = request.files.get("kmz")
        sheet_file = request.files.get("cutsheet")

        if not kmz_file or not kmz_file.filename.lower().endswith(".kmz"):
            flash("Please upload a .kmz file.")
            return redirect(url_for("index"))
        if not sheet_file or not sheet_file.filename.lower().endswith(".xlsx"):
            flash("Please upload the cut sheet (.xlsx).")
            return redirect(url_for("index"))

        # Create an isolated temp job directory.
        job_id  = _make_job_id()
        job_dir = Path(tempfile.mkdtemp(prefix=f"pipeline_{job_id}_"))
        data_dir   = job_dir / "data"
        output_dir = job_dir / "output"
        data_dir.mkdir(); output_dir.mkdir()

        # Save required files.
        kmz_file.save(data_dir / kmz_file.filename)
        sheet_file.save(data_dir / sheet_file.filename)

        # Save optional files.
        haf_file = request.files.get("haf")
        if haf_file and haf_file.filename:
            haf_file.save(data_dir / haf_file.filename)

        # Register job.
        _JOBS[job_id] = {
            "queue":              queue.Queue(),
            "thread":             None,
            "checkpoint_event":   threading.Event(),
            "checkpoint_continue": True,
            "job_dir":            str(job_dir),
            "status":             "running",
            "outputs":            {},
        }

        # Launch background thread.
        t = threading.Thread(
            target=_run_job,
            args=(job_id, str(data_dir), str(output_dir)),
            daemon=True,
        )
        t.start()
        _JOBS[job_id]["thread"] = t

        return redirect(url_for("run_page", job_id=job_id))

    # -----------------------------------------------------------------------
    # Live run page
    # -----------------------------------------------------------------------
    @app.route("/run/<job_id>")
    def run_page(job_id):
        if job_id not in _JOBS:
            flash("Job not found.")
            return redirect(url_for("index"))
        return render_template("run.html", job_id=job_id)

    # -----------------------------------------------------------------------
    # SSE endpoint — streams log lines to the browser
    # -----------------------------------------------------------------------
    @app.route("/stream/<job_id>")
    def stream(job_id):
        if job_id not in _JOBS:
            return Response("data: {}\n\n", mimetype="text/event-stream")

        job = _JOBS[job_id]

        def generate():
            while True:
                try:
                    msg = job["queue"].get(timeout=30)
                except queue.Empty:
                    yield "data: {\"type\": \"heartbeat\"}\n\n"
                    continue

                yield f"data: {json.dumps(msg)}\n\n"

                if msg["type"] == "checkpoint":
                    # Pause the stream and wait — the browser will redirect to /checkpoint/
                    # The pipeline thread is also waiting. We just stop sending new events.
                    break

                if msg["type"] == "done":
                    break

        return Response(generate(), mimetype="text/event-stream",
                        headers={"Cache-Control": "no-cache",
                                 "X-Accel-Buffering": "no"})

    # -----------------------------------------------------------------------
    # Checkpoint page
    # -----------------------------------------------------------------------
    @app.route("/checkpoint/<job_id>", methods=["GET"])
    def checkpoint_page(job_id):
        if job_id not in _JOBS:
            flash("Job not found.")
            return redirect(url_for("index"))
        job = _JOBS[job_id]
        # Look for the Colored Connections Table in the output dir.
        job_dir    = Path(job["job_dir"])
        conn_table = job_dir / "output" / "Colored_Connections_Table.xlsx"
        return render_template("checkpoint.html", job_id=job_id,
                               has_conn_table=conn_table.exists())

    @app.route("/checkpoint/<job_id>/continue", methods=["POST"])
    def checkpoint_continue(job_id):
        if job_id in _JOBS:
            _JOBS[job_id]["checkpoint_continue"] = True
            _JOBS[job_id]["checkpoint_event"].set()
        return redirect(url_for("run_page", job_id=job_id))

    @app.route("/checkpoint/<job_id>/stop", methods=["POST"])
    def checkpoint_stop(job_id):
        if job_id in _JOBS:
            _JOBS[job_id]["checkpoint_continue"] = False
            _JOBS[job_id]["checkpoint_event"].set()
        return redirect(url_for("complete_page", job_id=job_id))

    # -----------------------------------------------------------------------
    # GeoJSON API — cable network map data for the checkpoint Leaflet map
    # -----------------------------------------------------------------------
    @app.route("/geojson/<job_id>")
    def geojson(job_id):
        if job_id not in _JOBS:
            return jsonify({"type": "FeatureCollection", "features": []})
        try:
            data = build_geojson(_JOBS[job_id]["job_dir"])
        except Exception as exc:
            data = {"type": "FeatureCollection", "features": [],
                    "error": str(exc)}
        return jsonify(data)

    # -----------------------------------------------------------------------
    # Direction override — update cable fill in the Colored Connections Table
    # -----------------------------------------------------------------------
    _DIR_TO_HEX = {
        "NORTH": "FFA500", "SOUTH": "8B4513", "EAST": "008000",
        "WEST":  "708090", "OLT":   "C5D9B5", "MST":  "FF0000",
    }
    _DIR_LABELS = {
        "NORTH": "North", "SOUTH": "South", "EAST": "East",
        "WEST":  "West",  "OLT":   "OLT",   "MST":  "MST",
    }

    @app.route("/override/<job_id>", methods=["POST"])
    def override_direction(job_id):
        if job_id not in _JOBS:
            return jsonify({"ok": False, "error": "Job not found"}), 404
        data      = request.get_json(force=True) or {}
        cable     = str(data.get("cable", "")).strip()
        direction = str(data.get("direction", "")).upper()
        if not cable or direction not in _DIR_TO_HEX:
            return jsonify({"ok": False, "error": "Invalid cable or direction"}), 400

        hex6      = _DIR_TO_HEX[direction]
        new_fill  = PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")
        css_color = f"#{hex6}"

        colored_path = Path(_JOBS[job_id]["job_dir"]) / "output" / "Colored_Connections_Table.xlsx"
        if not colored_path.exists():
            return jsonify({"ok": False, "error": "Connections table not found"}), 404

        wb = openpyxl.load_workbook(str(colored_path))
        ws = wb.active
        changed = 0
        for row in ws.iter_rows(min_row=2):
            for cell in row[3:]:   # skip Location, Lat, Lon columns
                if str(cell.value or "").strip() == cable:
                    cell.fill = new_fill
                    changed  += 1
        wb.save(str(colored_path))

        return jsonify({"ok": True, "changed": changed,
                        "color": css_color, "direction": _DIR_LABELS[direction]})

    # -----------------------------------------------------------------------
    # Cable size rename — updates CT suffix in both xlsx tables
    # -----------------------------------------------------------------------
    @app.route("/resize/<job_id>", methods=["POST"])
    def resize_cable(job_id):
        if job_id not in _JOBS:
            return jsonify({"ok": False, "error": "Job not found"}), 404
        data     = request.get_json(force=True) or {}
        cable    = str(data.get("cable", "")).strip()
        new_size = str(data.get("size",  "")).strip().upper()
        if not cable or not re.match(r'^\d{2,3}CT$', new_size):
            return jsonify({"ok": False, "error": "Invalid cable or size"}), 400

        new_cable = _rename_cable_size(cable, new_size)
        if new_cable == cable:
            return jsonify({"ok": True, "old_cable": cable, "new_cable": cable, "changed": 0})

        job_dir = Path(_JOBS[job_id]["job_dir"])
        changed_total = 0
        for fpath in [
            job_dir / "data"   / "Connections_Table.xlsx",
            job_dir / "output" / "Colored_Connections_Table.xlsx",
        ]:
            if not fpath.exists():
                continue
            wb = openpyxl.load_workbook(str(fpath))
            changed = 0
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if str(cell.value or "").strip() == cable:
                            cell.value = new_cable
                            changed += 1
            if changed:
                wb.save(str(fpath))
                changed_total += changed

        return jsonify({"ok": True, "old_cable": cable, "new_cable": new_cable,
                        "changed": changed_total})

    # -----------------------------------------------------------------------
    # Completion page
    # -----------------------------------------------------------------------
    @app.route("/complete/<job_id>")
    def complete_page(job_id):
        if job_id not in _JOBS:
            flash("Job not found.")
            return redirect(url_for("index"))
        job = _JOBS[job_id]
        return render_template("complete.html", job_id=job_id,
                               status=job["status"],
                               outputs=list(job["outputs"].keys()))

    # -----------------------------------------------------------------------
    # File downloads
    # -----------------------------------------------------------------------
    @app.route("/download/<job_id>/<filename>")
    def download(job_id, filename):
        if job_id not in _JOBS:
            flash("Job not found.")
            return redirect(url_for("index"))
        job     = _JOBS[job_id]
        outputs = job.get("outputs", {})
        if filename not in outputs:
            flash(f"{filename} not found.")
            return redirect(url_for("complete_page", job_id=job_id))
        return send_file(outputs[filename], as_attachment=True,
                         download_name=filename)

    @app.route("/download/<job_id>/connections")
    def download_connections(job_id):
        if job_id not in _JOBS:
            return redirect(url_for("index"))
        fpath = Path(_JOBS[job_id]["job_dir"]) / "output" / "Colored_Connections_Table.xlsx"
        if not fpath.exists():
            flash("Connections table not available yet.")
            return redirect(url_for("checkpoint_page", job_id=job_id))
        return send_file(str(fpath), as_attachment=True,
                         download_name="Colored_Connections_Table.xlsx")

    return app
