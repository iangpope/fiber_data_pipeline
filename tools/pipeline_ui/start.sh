#!/usr/bin/env bash
# start.sh -- Start the Pipeline Web UI
# Usage: bash tools/pipeline_ui/start.sh

TOOL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TOOLS_DIR="$(cd "$TOOL_DIR/.." && pwd)"
VENV="$TOOLS_DIR/.venv/bin/python"

echo ""
echo " Fiber Pipeline Web UI"
echo " ─────────────────────────────────────────────"
echo " Open http://localhost:5002 in your browser"
echo " Press Ctrl+C to stop"
echo ""

cd "$TOOL_DIR"
exec "$VENV" run.py
