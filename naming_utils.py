"""
naming_utils.py -- Centralized location naming and classification utilities.

All pipeline scripts import from this module to avoid duplicating parsing and
classification logic. Two naming conventions are supported:

  New format:    MS90E_FT_001  (tap),  MS90E_SE_001  (splice enclosure)
  Legacy format: MICMS02S007   (MIC prefix, type letter, number suffix)

Public API
----------
  parse_location_id(name)           -> dict of parsed fields
  classify_location(name, df=None)  -> dict with kind/subtype/reasons
  is_tap(name, df=None)             -> bool
  is_splitter_enclosure(name)       -> bool
  is_distribution_enclosure(name)   -> bool
  is_olt_token(token)               -> bool
  classify_sheet_name(name)         -> 'T' / 'S' / 'D' / None
  is_location_sheet(name)           -> bool
  find_header_row(ws, ...)          -> int | None
  find_col_by_header(ws, ...)       -> int | None
  find_all_cols_by_header(ws, ...)  -> list[int]
  sheet_has_optical_splitters(ws)   -> bool
  optical_splitters_row(ws)         -> int | None
  safe_fill_hex(cell)               -> str | None
"""

import re


# ---------------------------------------------------------------------------
# Regular expressions for parsing location names
# ---------------------------------------------------------------------------

# New naming convention: e.g. RC73E_FT_001 or RC73E_SE_001
# Groups: olt (site prefix ending in E), class (FT or SE), num (3-digit number)
NEW_RX = re.compile(
    r'^(?P<olt>[A-Z0-9]+E)_(?P<class>FT|SE)_(?P<num>\d{3})$',
    re.IGNORECASE,
)

# Legacy naming convention: varying formats, the common pattern is an uppercase
# prefix ending with an optional type letter (S or D) and a numeric suffix.
# Examples: MICMS02S007, MICRCT093
LEGACY_SIMPLE_RX = re.compile(
    r'^(?P<prefix>[A-Z]{2}[A-Z0-9]+?)(?P<type>S|D)?(?P<num>\d{3,4})$',
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Location ID parsing
# ---------------------------------------------------------------------------

def parse_location_id(name: str) -> dict:
    """
    Parse a location name string and return a dict of its components.

    For the new naming scheme (RC73E_FT_001):
        scheme  : 'new'
        olt     : site prefix (e.g. 'RC73E')
        class   : 'FT' (tap) or 'SE' (splice enclosure)
        number  : integer port/enclosure number

    For legacy naming (MICMS02S007):
        scheme      : 'legacy'
        olt         : the prefix portion
        class       : None (not encoded in name)
        legacy_hint : 'S', 'D', or '' (type suffix when present)
        number      : integer suffix

    Returns {'scheme': 'unknown', ...} if the name does not match either pattern.
    """
    raw = name.strip() if name else ""
    if not raw:
        return {"scheme": "unknown", "raw": name, "olt": None, "class": None, "number": None}

    m = NEW_RX.match(raw)
    if m:
        d = m.groupdict()
        return {
            "scheme": "new",
            "raw":    raw,
            "olt":    d["olt"].upper(),
            "class":  d["class"].upper(),   # 'FT' or 'SE'
            "number": int(d["num"]),
        }

    m2 = LEGACY_SIMPLE_RX.match(raw)
    if m2:
        d = m2.groupdict()
        legacy_type = (d.get("type") or "").upper()   # 'S', 'D', or ''
        return {
            "scheme":      "legacy",
            "raw":         raw,
            "olt":         d["prefix"].upper(),
            "class":       None,          # FT/SE distinction not in name
            "legacy_hint": legacy_type,   # type suffix if present
            "number":      int(d["num"]),
        }

    return {"scheme": "unknown", "raw": raw, "olt": None, "class": None, "number": None}


# ---------------------------------------------------------------------------
# Sheet content heuristics for classification
# ---------------------------------------------------------------------------

# Pattern to detect splitter ratio strings (e.g. 1x2, 1x32) in sheet data.
RATIO_RX = re.compile(r'\b1x(2|4|8|16|32|64)\b', re.IGNORECASE)

# Keywords that indicate a splitter (rather than a distribution) enclosure.
SPLITTER_WORDS = re.compile(r'\b(split|splitter|mux|demux)\b', re.IGNORECASE)


def _df_has_splitter_signals(df) -> tuple:
    """
    Scan a pandas DataFrame (a small preview of a sheet) for indicators that
    the sheet represents a splitter enclosure rather than a distribution enclosure.

    Checks:
      1. Column headers contain 'OPTICAL SPLITTERS'
      2. First two columns contain splitter ratio strings (1x2, 1x32, etc.)
      3. First two columns contain splitter keywords (splitter, mux, demux)

    Returns (found: bool, reason: str).
    """
    try:
        # Check column header row for the OPTICAL SPLITTERS marker.
        col_names  = [str(c) for c in df.columns]
        header_str = " ".join(col_names)
        if "OPTICAL SPLITTERS" in header_str.upper():
            return True, "Header mentions OPTICAL SPLITTERS"

        # Sample the first 150 rows of the first two data columns to keep this fast.
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


# ---------------------------------------------------------------------------
# Location classification
# ---------------------------------------------------------------------------

def classify_location(name: str, df=None) -> dict:
    """
    Classify a location name into kind/subtype/reasons.

    Possible kinds:
      'tap'                  -- FT enclosure (tap/OTE)
      'splice_enclosure'     -- SE enclosure (splitter or distribution)
      'splice_enclosure_or_tap' -- ambiguous legacy name
      'unknown'              -- unrecognized format

    Subtypes for splice_enclosure:
      'splitter'             -- contains optical splitters or MUX/DEMUX
      'distribution'         -- cable pass-through only
      'unknown'              -- SE without sheet content to distinguish

    If df (a pandas DataFrame of the sheet's content) is provided, splitter
    keyword heuristics are applied to resolve the splitter vs. distribution
    ambiguity for SE enclosures.
    """
    meta   = parse_location_id(name)
    scheme = meta.get("scheme")
    cls    = (meta.get("class") or "").upper()

    if scheme == "new":
        if cls == "FT":
            return {"kind": "tap", "subtype": "FT", "reasons": ["New scheme FT label"]}
        if cls == "SE":
            if df is not None:
                has_split, why = _df_has_splitter_signals(df)
                if has_split:
                    return {"kind": "splice_enclosure", "subtype": "splitter",     "reasons": [why]}
                else:
                    return {"kind": "splice_enclosure", "subtype": "distribution", "reasons": [why]}
            # Without sheet data, cannot determine splitter vs. distribution.
            return {"kind": "splice_enclosure", "subtype": "unknown", "reasons": ["SE without sheet context"]}

    if scheme == "legacy":
        # Legacy names don't encode FT/SE directly; use sheet content if available.
        if df is not None:
            has_split, why = _df_has_splitter_signals(df)
            if has_split:
                return {"kind": "splice_enclosure",         "subtype": "splitter",         "reasons": [f"Legacy + {why}"]}
            else:
                return {"kind": "splice_enclosure_or_tap",  "subtype": "distribution_or_tap", "reasons": [f"Legacy + {why}"]}
        return {"kind": "unknown", "subtype": "unknown", "reasons": ["Legacy name only"]}

    return {"kind": "unknown", "subtype": "unknown", "reasons": ["Unrecognized name format"]}


# ---------------------------------------------------------------------------
# Convenience boolean classifiers
# ---------------------------------------------------------------------------

def is_tap(name: str, df=None) -> bool:
    """Return True if the location name represents a tap (FT/OTE) enclosure."""
    c = classify_location(name, df=df)
    return c["kind"] == "tap"


def is_splitter_enclosure(name: str, df=None) -> bool:
    """Return True if the location is a splice enclosure containing optical splitters."""
    c = classify_location(name, df=df)
    return c["kind"] == "splice_enclosure" and c["subtype"] == "splitter"


def is_distribution_enclosure(name: str, df=None) -> bool:
    """Return True if the location is a distribution (pass-through) splice enclosure."""
    c = classify_location(name, df=df)
    return c["kind"] == "splice_enclosure" and c["subtype"] == "distribution"


# ---------------------------------------------------------------------------
# OLT token recognition
# ---------------------------------------------------------------------------

# Bare OLT site token pattern: 2-10 alphanumeric characters ending with 'E'.
# Examples: RC73E, MS33E, PT21E, SH20E
OLT_TOKEN_RX = re.compile(r"^[A-Z0-9]{2,10}E$", re.IGNORECASE)


def is_olt_token(token: str) -> bool:
    """
    Return True if the string looks like a bare OLT/node site identifier.
    These are site codes like RC73E that identify head-end racks, not
    individual splice locations.
    """
    if not token:
        return False
    return bool(OLT_TOKEN_RX.match(str(token).strip()))

# Back-compat alias used by some scripts that import 'is_olt'.
is_olt = is_olt_token


# ---------------------------------------------------------------------------
# Lightweight sheet-name classification (for formatting scripts)
# ---------------------------------------------------------------------------

def classify_sheet_name(name: str) -> str | None:
    """
    Return a single-character classification for a location sheet name:
      'T'  -- tap enclosure (FT)
      'S'  -- splice enclosure (SE, splitter assumed without sheet data)
      'D'  -- distribution enclosure
      None -- unknown or unrecognized

    This is a quick classification based on the name alone, without sheet data.
    It is used by formatting scripts that don't have access to the sheet content
    at classification time.
    """
    info   = parse_location_id(name or "")
    scheme = info.get("scheme")

    if scheme == "new":
        cls = info.get("class")
        if cls == "FT":
            return "T"
        if cls == "SE":
            # Cannot determine splitter vs. distribution without sheet data;
            # default to splitter enclosure ('S') as the more common case.
            return "S"

    if scheme == "legacy":
        hint = info.get("legacy_hint")
        if hint == "S":
            return "S"
        if hint == "D":
            return "D"
        # Legacy tap enclosures tend to have 4-digit suffixes; check via heuristic.
        if is_tap(name):
            return "T"

    return None


# ---------------------------------------------------------------------------
# Shared worksheet utilities -- header-based column and row detection
# ---------------------------------------------------------------------------

def find_header_row(ws, label: str = "CONNECTION", max_scan_rows: int = 120) -> int | None:
    """
    Scan the worksheet for the first row that contains a cell matching label
    (case-insensitive). Returns the 1-based row number, or None if not found.

    This is used to locate the SHEATHS data header row without relying on a
    hardcoded row number, making the pipeline resilient to sheets with varying
    amounts of metadata above the data table.
    """
    label_up = label.strip().upper()
    max_r    = min(ws.max_row, max_scan_rows)
    max_c    = min(ws.max_column, 120)
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
    Find the column in header_row whose cell value matches label (case-insensitive).

    Parameters:
        after_col : if set, only columns strictly greater than this index are
                    considered. Useful for finding a second occurrence of a
                    repeated header (e.g. finding the SHEATH UUID column that
                    appears to the right of the CONNECTION column).

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
    """
    Return all column indices in header_row where the cell value matches label
    (case-insensitive). Used when a column name appears more than once in the
    header row (e.g. SHEATH UUID appears in both the main and splitter tables).
    """
    label_up = label.strip().upper()
    if max_col is None:
        max_col = ws.max_column
    return [
        c for c in range(1, max_col + 1)
        if isinstance(ws.cell(header_row, c).value, str)
        and ws.cell(header_row, c).value.strip().upper() == label_up
    ]


# ---------------------------------------------------------------------------
# Sheet content detection helpers
# ---------------------------------------------------------------------------

def sheet_has_optical_splitters(ws, max_scan_rows: int = 400) -> bool:
    """
    Return True if column A of the worksheet contains a cell with the text
    'OPTICAL SPLITTERS' (case-insensitive).

    This is a content-based check, not a name-based check, so it correctly
    identifies SE sheets regardless of the naming convention in use.
    """
    max_r = min(ws.max_row, max_scan_rows)
    for r in range(1, max_r + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "OPTICAL SPLITTERS" in v.upper():
            return True
    return False


def optical_splitters_row(ws, max_scan_rows: int = 400) -> int | None:
    """
    Return the 1-based row number of the 'OPTICAL SPLITTERS' section header
    in column A, or None if the section is not present on this sheet.
    """
    max_r = min(ws.max_row, max_scan_rows)
    for r in range(1, max_r + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and "OPTICAL SPLITTERS" in v.upper():
            return r
    return None


def is_location_sheet(name: str) -> bool:
    """
    Return True if the sheet name represents a splice or tap location that
    should be processed by the pipeline.

    Sheets named 'Index', 'Legend', 'Notes', 'Sheet1', or starting with
    'Tap Report' are skipped because they are administrative, not splice data.
    """
    n = name.strip().lower()
    if n in {"index", "legend", "notes", "sheet1"}:
        return False
    if n.startswith("tap report"):
        return False
    return True


# ---------------------------------------------------------------------------
# Safe cell fill color extraction
# ---------------------------------------------------------------------------

def safe_fill_hex(cell) -> str | None:
    """
    Return the 6-character RGB hex string of a cell's solid fill color,
    or None if the cell has no fill, a non-solid fill, or a non-RGB color type.

    openpyxl can return indexed or theme-based Color objects for cells whose
    fill was set by Excel's conditional formatting or theme system. Accessing
    .rgb on those types raises an exception or returns a nonsense value. This
    function guards against all such cases and always returns a plain string
    or None.
    """
    fill = getattr(cell, "fill", None)
    if fill is None or getattr(fill, "fill_type", None) != "solid":
        return None

    sc = getattr(fill, "start_color", None)
    if sc is None:
        return None

    # Only process RGB-type colors; indexed and theme types are not usable.
    if getattr(sc, "type", None) != "rgb":
        return None

    rgb = getattr(sc, "rgb", None)
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None

    # openpyxl may return 8-char ARGB (e.g. 'FFRRGGBB'); take the last 6.
    return rgb[-6:].upper()
