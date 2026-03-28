"""
app.py -- Flask web application for the Cable Resize Tool.

Routes:
    GET  /              -- Landing page with file upload
    POST /review        -- Parse uploaded workbook, show resize table
    POST /apply         -- Apply selections, return modified workbook download
"""

from __future__ import annotations

import io
import os
import sys
from pathlib import Path

from flask import (
    Flask, request, render_template, send_file, session, redirect, url_for, flash
)
import openpyxl

# Add the pipeline root to sys.path so resize_logic can find config/naming_utils.
_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from resize_logic import scan_sheaths, apply_all_resizes

try:
    from config import VALID_CT_SIZES
except ImportError:
    VALID_CT_SIZES = [12, 24, 48, 96, 144]


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.urandom(32)

    # Store uploaded workbook bytes in the server-side filesystem session.
    app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB max upload


    # -----------------------------------------------------------------------
    # Landing page
    # -----------------------------------------------------------------------
    @app.route("/", methods=["GET"])
    def index():
        return render_template("index.html")


    # -----------------------------------------------------------------------
    # Review page: parse workbook and show sheath table
    # -----------------------------------------------------------------------
    @app.route("/review", methods=["POST"])
    def review():
        if "workbook" not in request.files:
            flash("No file uploaded.")
            return redirect(url_for("index"))

        file = request.files["workbook"]
        if not file.filename.lower().endswith(".xlsx"):
            flash("Please upload an .xlsx file.")
            return redirect(url_for("index"))

        wb_bytes = file.read()

        try:
            wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
        except Exception as e:
            flash(f"Could not open workbook: {e}")
            return redirect(url_for("index"))

        records = scan_sheaths(wb)

        if not records:
            flash("No sheath blocks found. Make sure this is a completed asbuilt workbook.")
            return redirect(url_for("index"))

        # Store raw bytes in the session filesystem for /apply to use.
        session["wb_bytes"] = wb_bytes
        session["filename"] = file.filename

        # Group records by sheet for the template.
        by_sheet: dict[str, list] = {}
        for rec in records:
            by_sheet.setdefault(rec["sheet"], []).append(rec)

        return render_template(
            "review.html",
            by_sheet=by_sheet,
            ct_sizes=VALID_CT_SIZES,
            filename=file.filename,
        )


    # -----------------------------------------------------------------------
    # Apply resizes and return the modified workbook as a download
    # -----------------------------------------------------------------------
    @app.route("/apply", methods=["POST"])
    def apply():
        wb_bytes = session.get("wb_bytes")
        filename  = session.get("filename", "resized.xlsx")

        if not wb_bytes:
            flash("Session expired. Please re-upload your workbook.")
            return redirect(url_for("index"))

        wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))

        # Build resize map from form: field names are "resize_<uuid>"
        resize_map = {}
        for key, val in request.form.items():
            if key.startswith("resize_") and val:
                uuid = key[len("resize_"):]
                try:
                    new_ct = int(val)
                    resize_map[uuid] = new_ct
                except ValueError:
                    pass

        errors = apply_all_resizes(wb, resize_map)

        if errors:
            # Return to review with error messages shown.
            flash("Some resizes could not be applied:\n" + "\n".join(errors))
            return redirect(url_for("index"))

        # Save modified workbook to an in-memory buffer and send as download.
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        stem     = Path(filename).stem
        out_name = f"{stem}_resized.xlsx"

        return send_file(
            buf,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    return app
