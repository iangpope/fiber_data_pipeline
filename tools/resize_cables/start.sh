#!/usr/bin/env bash
# start.sh -- Start the Cable Resize Tool web interface
# Usage: bash tools/resize_cables/start.sh

TOOL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TOOLS_DIR="$(cd "$TOOL_DIR/.." && pwd)"
VENV="$TOOLS_DIR/.venv/bin/python"

echo ""
echo " Cable Resize Tool"
echo " ─────────────────────────────────────────────"
echo " Open http://localhost:5001 in your browser"
echo " Press Ctrl+C to stop"
echo ""

cd "$TOOL_DIR"
exec "$VENV" run.py
