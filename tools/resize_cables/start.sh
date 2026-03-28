#!/usr/bin/env bash
# start.sh -- Start the Cable Resize Tool web interface
# Usage: bash tools/resize_cables/start.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$SCRIPT_DIR/.venv/bin/python"

echo ""
echo " Cable Resize Tool"
echo " ─────────────────────────────────────────────"
echo " Open http://localhost:5001 in your browser"
echo " Press Ctrl+C to stop"
echo ""

cd "$SCRIPT_DIR"
exec "$VENV" run.py
