"""
naming_utils.py — centralize FTTH location naming + classification logic (legacy + new scheme).

Supports:
- Legacy names like: MICMS02S007, "48CT A TO B" sheaths, etc.
- New names like:    MS90E_FT_001 (tap), MS90E_SE_001 (splice enclosure: splitter or distribution)

API:
    parse_location_id(name) -> dict with keys: olt, class, number, scheme, raw
    classify_location(name, df=None) -> dict with keys: kind, subtype, reasons
    is_tap(name, df=None) -> bool
    is_splitter_enclosure(name, df=None) -> bool
    is_distribution_enclosure(name, df=None) -> bool

Heuristics for splitter-vs-distribution (when df is provided for that sheet):
- Looks for "OPTICAL SPLITTERS" in any cell of col A (case-insensitive)
- Looks for ratio-like strings (1x2, 1x4, 1x8, …) in column B or near header
- Looks for keywords ["splitter", "mux", "demux"]
- If none of the above, default subtype='distribution' for SE
"""

import re

NEW_RX = re.compile(r'^(?P<olt>[A-Z0-9]+E)_(?P<class>FT|SE)_(?P<num>\d{3})$', re.IGNORECASE)

# Legacy formats vary a lot; capture common patterns
# Examples seen: MICMS02S007  (MI + CMS02 + S + 007)
LEGACY_SIMPLE_RX = re.compile(r'^(?P<prefix>[A-Z]{2}[A-Z0-9]+?)(?P<type>S|D)?(?P<num>\d{3,4})$', re.IGNORECASE)

def parse_location_id(name: str):
    raw = name.strip() if name else ""
    if not raw:
        return {"scheme": "unknown", "raw": name, "olt": None, "class": None, "number": None}

    m = NEW_RX.match(raw)
    if m:
        d = m.groupdict()
        return {
            "scheme": "new",
            "raw": raw,
            "olt": d["olt"].upper(),
            "class": d["class"].upper(),   # FT or SE
            "number": int(d["num"]),
        }

    m2 = LEGACY_SIMPLE_RX.match(raw)
    if m2:
        d = m2.groupdict()
        # Legacy doesn't encode FT/SE directly
        legacy_type = (d.get("type") or "").upper()  # S or D sometimes
        return {
            "scheme": "legacy",
            "raw": raw,
            "olt": d["prefix"].upper(),
            "class": None,                # unknown from name alone
            "legacy_hint": legacy_type,   # S/D if present
            "number": int(d["num"]),
        }

    # Unknown style, return safe default
    return {"scheme": "unknown", "raw": raw, "olt": None, "class": None, "number": None}

RATIO_RX = re.compile(r'\b1x(2|4|8|16|32|64)\b', re.IGNORECASE)
SPLITTER_WORDS = re.compile(r'\b(split|splitter|mux|demux)\b', re.IGNORECASE)

def _df_has_splitter_signals(df):
    try:
        # Check column A header area for "OPTICAL SPLITTERS"
        col_names = [str(c) for c in df.columns]
        header_str = " ".join(col_names)
        if "OPTICAL SPLITTERS" in header_str.upper():
            return True, "Header mentions OPTICAL SPLITTERS"

        # Scan first ~1000 cells of first two columns for ratio or splitter words
        # (kept small to be fast; callers should pass a small preview df)
        sample_cols = df.columns[:2]
        sample = df[list(sample_cols)].astype(str).head(150).fillna("")
        joined = " ".join([" ".join(row) for _, row in sample.iterrows()])
        if RATIO_RX.search(joined):
            return True, "Found splitter ratio like 1x4/1x8"
        if SPLITTER_WORDS.search(joined):
            return True, "Found 'splitter'/'mux'/'demux' keywords"
        return False, "No splitter signals found"
    except Exception as e:
        return False, f"Error scanning df: {e}"

def classify_location(name: str, df=None):
    meta = parse_location_id(name)
    scheme = meta.get("scheme")
    cls = (meta.get("class") or "").upper()

    if scheme == "new":
        if cls == "FT":
            return {"kind": "tap", "subtype": "FT", "reasons": ["New scheme FT label"]}
        if cls == "SE":
            if df is not None:
                has_split, why = _df_has_splitter_signals(df)
                if has_split:
                    return {"kind": "splice_enclosure", "subtype": "splitter", "reasons": [why]}
                else:
                    return {"kind": "splice_enclosure", "subtype": "distribution", "reasons": [why]}
            # Without df, default unknown subtype
            return {"kind": "splice_enclosure", "subtype": "unknown", "reasons": ["SE without sheet context"]}

    if scheme == "legacy":
        # Can't tell from name alone; use df if available
        if df is not None:
            has_split, why = _df_has_splitter_signals(df)
            if has_split:
                return {"kind": "splice_enclosure", "subtype": "splitter", "reasons": [f"Legacy + {why}"]}
            else:
                return {"kind": "splice_enclosure_or_tap", "subtype": "distribution_or_tap", "reasons": [f"Legacy + {why}"]}
        # fallback
        return {"kind": "unknown", "subtype": "unknown", "reasons": ["Legacy name only"]}

    # Unknown naming scheme
    return {"kind": "unknown", "subtype": "unknown", "reasons": ["Unrecognized name format"]}

def is_tap(name: str, df=None) -> bool:
    c = classify_location(name, df=df)
    if c["kind"] == "tap":
        return True
    # Back-compat geo/graph rule: callers may add distance-based tap detection
    return False

def is_splitter_enclosure(name: str, df=None) -> bool:
    c = classify_location(name, df=df)
    return c["kind"] == "splice_enclosure" and c["subtype"] == "splitter"

def is_distribution_enclosure(name: str, df=None) -> bool:
    c = classify_location(name, df=df)
    return c["kind"] == "splice_enclosure" and c["subtype"] == "distribution"


# -----------------------------
# Additions for pipeline scripts
# -----------------------------
# Bare OLT token like "RC077E", "MS33E", "PT21E", "SH20E"
OLT_TOKEN_RX = re.compile(r"^[A-Z0-9]{2,10}E$", re.IGNORECASE)

def is_olt_token(token: str) -> bool:
    """Return True if token looks like a bare OLT/node id (e.g., RC077E, MS33E)."""
    if not token:
        return False
    return bool(OLT_TOKEN_RX.match(str(token).strip()))

# Back-compat alias (some scripts may import this name)
is_olt = is_olt_token

def classify_sheet_name(name: str) -> str | None:
    """Lightweight classification used by formatting scripts.
    Returns:
        'T' tap, 'S' splitter enclosure, 'D' distribution enclosure, or None unknown.
    """
    info = parse_location_id(name or "")
    scheme = info.get("scheme")
    if scheme == "new":
        cls = info.get("class")
        if cls == "FT":
            return "T"
        if cls == "SE":
            # Need sheet content to decide split vs distro; default to enclosure (S)
            return "S"
    if scheme == "legacy":
        hint = info.get("legacy_hint")
        if hint == "S":
            return "S"
        if hint == "D":
            return "D"
        # Legacy taps are the 4-digit suffix case in many of your builds (e.g., MICMS670023)
        # parse_location_id returns class None; classify_location handles this via heuristics.
        if is_tap(name):
            return "T"
    return None


# ---------------------------------------------------------------------------
# Shared worksheet utilities — header-based column detection
# ---------------------------------------------------------------------------

def find_header_row(ws, label: str = "CONNECTION", max_scan_rows: int = 120) -> int | None:
    """
    Scan the worksheet for the row containing *label* (case-insensitive).
    Returns the 1-based row number, or None if not found.
    """
    label_up = label.strip().upper()
    max_r = min(ws.max_row, max_scan_rows)
    max_c = min(ws.max_column, 120)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().upper() == label_up:
                return r
    return None


def find_col_by_header(ws, header_row: int, label: str,
                       min_col: int = 1, max_col: int | None = None,
                       after_col: int | None = None) -> int | None:
    """
    Find the column in *header_row* whose value matches *label* (case-insensitive).

    Args:
        after_col: if set, only consider columns strictly greater than this index.

    Returns the 1-based column number, or None if not found.
    """
    label_up = label.strip().upper()
    if max_col is None:
        max_col = ws.max_column
    for c in range(min_col, max_col + 1):
        if after_col is not None and c <= after_col:
            continue
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip().upper() == label_up:
            return c
    return None


def find_all_cols_by_header(ws, header_row: int, label: str,
                             max_col: int | None = None) -> list[int]:
    """Return all column indices in *header_row* matching *label*."""
    label_up = label.strip().upper()
    if max_col is None:
        max_col = ws.max_column
    return [
        c for c in range(1, max_col + 1)
        if isinstance(ws.cell(header_row, c).value, str)
        and ws.cell(header_row, c).value.strip().upper() == label_up
    ]


def sheet_has_optical_splitters(ws, max_scan_rows: int = 400) -> bool:
    """
    Return True if column A of *ws* contains a cell with 'OPTICAL SPLITTERS'
    (case-insensitive). Content-based — works regardless of sheet name format.
    """
    max_r = min(ws.max_row, max_scan_rows)
    for r in range(1, max_r + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "OPTICAL SPLITTERS" in v.upper():
            return True
    return False


def optical_splitters_row(ws, max_scan_rows: int = 400) -> int | None:
    """
    Return the 1-based row number of the 'OPTICAL SPLITTERS' header row,
    or None if not present.
    """
    max_r = min(ws.max_row, max_scan_rows)
    for r in range(1, max_r + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "OPTICAL SPLITTERS" in v.upper():
            return r
    return None


def is_location_sheet(name: str) -> bool:
    """
    Return True if the sheet name represents a splice/tap location
    (i.e. should be processed by the pipeline).
    Excludes known non-location sheets.
    """
    n = name.strip().lower()
    if n in {"index", "legend", "notes", "sheet1"}:
        return False
    if n.startswith("tap report"):
        return False
    return True


def safe_fill_hex(cell) -> str | None:
    """
    Return the last-6-char RGB hex of a cell's solid fill,
    or None if the fill is absent, non-solid, or not an RGB type.
    Guards against indexed/theme colors that crash on .rgb access.
    """
    fill = getattr(cell, "fill", None)
    if fill is None or getattr(fill, "fill_type", None) != "solid":
        return None
    sc = getattr(fill, "start_color", None)
    if sc is None:
        return None
    if getattr(sc, "type", None) != "rgb":
        return None
    rgb = getattr(sc, "rgb", None)
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None
    return rgb[-6:].upper()
