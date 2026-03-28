"""
run.py -- Launch the Cable Resize Tool web interface.

Usage:
    cd tools/resize_cables
    python3 run.py

Then open http://localhost:5001 in your browser.
"""

from app import create_app

app = create_app()

if __name__ == "__main__":
    print("\n Cable Resize Tool")
    print(" ─────────────────────────────────────────────")
    print(" Open http://localhost:5001 in your browser")
    print(" Press Ctrl+C to stop\n")
    app.run(debug=False, port=5001, host="127.0.0.1")
