"""
config.py — Shared constants for the Fiber Data Pipeline.

All scripts import colors, paths, and skip lists from here
instead of defining their own.
"""

import os
from openpyxl.styles import PatternFill

# ---------------------------------------------------------------------------
# Directory paths
# ---------------------------------------------------------------------------
DATA_DIR   = "data"
OUTPUT_DIR = "output"

# ---------------------------------------------------------------------------
# Sheet names to always skip (lowercased for comparison)
# ---------------------------------------------------------------------------
SKIP_SHEETS = {"index", "legend", "notes", "sheet1", "tap report"}

# ---------------------------------------------------------------------------
# Color palette — hex strings (6 chars, no alpha)
# ---------------------------------------------------------------------------
COLOR = {
    # Directional sheath colors
    "NORTH":  "FFA500",   # orange
    "SOUTH":  "8B4513",   # brown
    "EAST":   "008000",   # green
    "WEST":   "708090",   # slate
    "MST":    "FF0000",   # red
    "OLT":    "C5D9B5",   # olive green

    # Splice / connection types
    "FUSION": "FFFF00",   # yellow
    "YELLOW": "FFFF00",   # alias

    # Splitter / MUX port types
    "COMMON": "ADD8E6",   # light blue
    "1X32":   "FFB6C1",   # light pink
    "1X2":    "DB7093",   # dark pink
    "MUX":    "FFDAB9",   # peach
    "DEMUX":  "FFA07A",   # salmon
}

# Map of color hex → directional label (used by step 2 and step 5)
COLOR_TO_DIRECTION = {
    COLOR["NORTH"]: "North",
    COLOR["SOUTH"]: "South",
    COLOR["EAST"]:  "East",
    COLOR["WEST"]:  "West",
    COLOR["OLT"]:   "OLT",
    COLOR["MST"]:   "MST",
}

# ---------------------------------------------------------------------------
# Pre-built PatternFill objects keyed by hex — avoids repeated construction
# ---------------------------------------------------------------------------
def _fill(hex6: str) -> PatternFill:
    return PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")

FILLS = {hex6: _fill(hex6) for hex6 in {
    "FFA500", "8B4513", "008000", "708090", "FF0000", "C5D9B5",
    "FFFF00", "ADD8E6", "FFB6C1", "DB7093", "FFDAB9", "FFA07A",
    "7FFF00", "FFF9DB",
}}

EMPTY_FILL = PatternFill()  # no fill

# ---------------------------------------------------------------------------
# Helper — get a PatternFill by COLOR key or raw hex
# ---------------------------------------------------------------------------
def get_fill(key_or_hex: str) -> PatternFill:
    """
    Return a PatternFill for a COLOR key (e.g. 'NORTH') or raw hex (e.g. 'FFA500').
    Creates one on-the-fly if not cached.
    """
    hex6 = COLOR.get(key_or_hex.upper(), key_or_hex).upper()[-6:]
    if hex6 not in FILLS:
        FILLS[hex6] = _fill(hex6)
    return FILLS[hex6]
