"""
run.py -- Launch the Pipeline Web UI.

Usage:
    cd tools/pipeline_ui
    python3 run.py

Then open http://localhost:5002 in your browser.
"""

from app import create_app

app = create_app()

if __name__ == "__main__":
    print("\n Fiber Pipeline Web UI")
    print(" ─────────────────────────────────────────────")
    print(" Open http://localhost:5002 in your browser")
    print(" Press Ctrl+C to stop\n")
    app.run(debug=False, port=5002, host="127.0.0.1", threaded=True)
