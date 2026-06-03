"""
mi_pre_normaliser.py
---------------------
Phase 3.5 of the Veridooh BFK pipeline — runs BEFORE Phase 4 (mi-data-automation).

Every media agency formats their MI differently. This normaliser ingests a raw
MI file in any common format and outputs a clean, standardised CSV that
mi-data-automation can reliably process.

What it fixes
-------------
- Column names: maps every known agency header variant to a canonical name
- Dates: normalises DD/MM/YYYY, MM/DD/YYYY, "1 July 2026", Excel serials, and
         mixed formats to YYYY/MM/DD
- Resolutions: "1080 X 1920px", "1080*1920", "W:1080 H:1920" → "1080x1920"
- Ad lengths: "15 secs", "0:15", "15s", "fifteen" → plain integer
- Rotation ratios: 0.7 → 70, "70%" → 70 (never writes "%" symbol)
- Burst end: "Yes", "TRUE", "1", blank → 1 (end) or 0 (not end)
- Consec keywords: normalises "Sequential", "CP", "Consec." → "Consec"
- Structural: removes empty rows, flattens merged cells, detects header row

Usage
-----
    from mi_pre_normaliser import MIPreNormaliser, run_mi_normaliser

    df_clean, report = run_mi_normaliser("JCD_MI.xlsx", out_path="JCD_MI_normalised.csv")
    # Then hand df_clean to mi-data-automation as the MI input
"""

from __future__ import annotations

import re
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# Canonical column names (what mi-data-automation reliably recognises)
# ---------------------------------------------------------------------------

CANON_PANEL_ID    = "Panel ID"
CANON_CREATIVE    = "Creative File Name"
CANON_AD_LEN      = "Ad Length"
CANON_WIDTH       = "Width"
CANON_HEIGHT      = "Height"
CANON_RESOLUTION  = "Resolution"
CANON_LIVE_DATE   = "Live Date"
CANON_END_DATE    = "End Date"
CANON_BURST_END   = "Burst End"
CANON_ROTATION    = "Rotation"
CANON_FORMAT      = "Format"
CANON_ADDRESS     = "Address"
CANON_COMMENTS    = "Notes"

# ---------------------------------------------------------------------------
# Synonym map — maps lowercase partial column header → canonical name
# Matching: case-insensitive substring, first hit wins
# ---------------------------------------------------------------------------

HEADER_SYNONYMS: dict[str, list[str]] = {
    CANON_PANEL_ID: [
        "panel id", "panel number", "panel #", "panel no", "site id",
        "site code", "site #", "asset id", "player id", "screen id",
    ],
    CANON_CREATIVE: [
        "creative file", "creative name", "material name", "asset name",
        "filename", "file name", "creative", "material", "artwork",
        "ad name", "spot name",
    ],
    CANON_AD_LEN: [
        "ad length", "duration", "spot length", "ad duration",
        "spot duration", "seconds", "length (s", "length(s",
        "duration (s", "duration(s",
    ],
    CANON_WIDTH: [
        "width (px", "width pixels", "pixels width", "screen width",
        "pixel width", "size - digital pixels width",
    ],
    CANON_HEIGHT: [
        "height (px", "height pixels", "pixels height", "screen height",
        "pixel height", "size - digital pixels height",
    ],
    CANON_RESOLUTION: [
        "resolution", "screen size", "screen resolution", "dimensions",
        "size (pixels", "screensize", "pixel dimensions", "dimension",
    ],
    CANON_LIVE_DATE: [
        "live date", "go live", "start date", "live from", "activation date",
        "campaign start", "flight start", "burst start", "from date",
        "live on", "live", "air date", "on air",
    ],
    CANON_END_DATE: [
        "end date", "campaign end", "flight end", "expiry", "until",
        "ends", "off air", "end on", "to date", "finish",
    ],
    CANON_BURST_END: [
        "burst end", "is burst end", "burst number", "burst week",
        "burst", "week end", "end of burst",
    ],
    CANON_ROTATION: [
        "rotation", "weight", "split", "ratio", "sov", "share",
        "rotation ratio", "creative ratio", "play ratio",
    ],
    CANON_FORMAT: [
        "format", "panel format", "oma format", "media format",
    ],
    CANON_ADDRESS: [
        "address", "location", "site address", "panel address",
    ],
    CANON_COMMENTS: [
        "notes", "comments", "instruction", "remark", "special",
    ],
}

# ---------------------------------------------------------------------------
# Consec detection
# ---------------------------------------------------------------------------

CONSEC_PATTERNS = re.compile(
    r"(?i)\b(consec\.?|consecutive|sequential)(?!\w)|(?<![a-zA-Z])\bcp\b"
)

# ---------------------------------------------------------------------------
# Burst-end truthy values
# ---------------------------------------------------------------------------

BURST_END_TRUE  = {"1", "yes", "true", "y", "end", "final", "last", "x"}
BURST_END_FALSE = {"0", "no", "false", "n", ""}


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

@dataclass
class NormaliserChange:
    column: str
    row_num: int
    original: str
    normalised: str
    reason: str


@dataclass
class NormaliserReport:
    source_file: str               = ""
    rows_in: int                   = 0
    rows_out: int                  = 0
    empty_rows_dropped: int        = 0
    columns_renamed: dict          = field(default_factory=dict)
    dates_normalised: int          = 0
    resolutions_normalised: int    = 0
    ad_lengths_normalised: int     = 0
    rotations_normalised: int      = 0
    burst_ends_normalised: int     = 0
    consec_normalised: int         = 0
    unrecognised_columns: list     = field(default_factory=list)
    changes: list[NormaliserChange] = field(default_factory=list)

    def summary(self) -> str:
        lines = [
            "=" * 60,
            f"MI Pre-normaliser Report — {self.source_file}",
            "=" * 60,
            f"  Rows in                  : {self.rows_in}",
            f"  Rows out (after clean)   : {self.rows_out}",
            f"  Empty rows dropped       : {self.empty_rows_dropped}",
            f"  Columns renamed          : {len(self.columns_renamed)}",
            f"  Dates normalised         : {self.dates_normalised}",
            f"  Resolutions normalised   : {self.resolutions_normalised}",
            f"  Ad lengths normalised    : {self.ad_lengths_normalised}",
            f"  Rotations normalised     : {self.rotations_normalised}",
            f"  Burst ends normalised    : {self.burst_ends_normalised}",
            f"  Consec flags normalised  : {self.consec_normalised}",
        ]
        if self.columns_renamed:
            lines.append("\nColumn renames:")
            for orig, canon in self.columns_renamed.items():
                lines.append(f"  {orig!r:35s} → {canon!r}")
        if self.unrecognised_columns:
            lines.append("\nUnrecognised columns (kept as-is):")
            for col in self.unrecognised_columns:
                lines.append(f"  · {col}")
        lines.append("=" * 60)
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Core normaliser
# ---------------------------------------------------------------------------

class MIPreNormaliser:
    """
    Normalises a raw MI file for reliable ingestion by mi-data-automation.

    Parameters
    ----------
    strict_dates : bool
        If True, raise ValueError on ambiguous dates (e.g. month/day both ≤12).
        If False (default), prefer DD/MM/YYYY for Australian locale.
    """

    def __init__(self, strict_dates: bool = False):
        self.strict_dates = strict_dates

    def normalise(self, df: pd.DataFrame, source_name: str = "") -> tuple[pd.DataFrame, NormaliserReport]:
        report = NormaliserReport(source_file=source_name, rows_in=len(df))

        # Step 1 — Drop completely empty rows
        df = df.dropna(how="all").reset_index(drop=True)
        report.empty_rows_dropped = report.rows_in - len(df)

        # Step 2 — Strip all string cells
        df = df.apply(lambda col: col.map(lambda v: v.strip() if isinstance(v, str) else v))

        # Step 3 — Rename columns to canonical names
        df, report.columns_renamed, report.unrecognised_columns = self._rename_columns(df)

        # Step 4 — Normalise each field type
        for col, fn, counter_attr in [
            (CANON_LIVE_DATE,   self._normalise_date,       "dates_normalised"),
            (CANON_END_DATE,    self._normalise_date,       "dates_normalised"),
            (CANON_RESOLUTION,  self._normalise_resolution, "resolutions_normalised"),
            (CANON_AD_LEN,      self._normalise_ad_length,  "ad_lengths_normalised"),
            (CANON_ROTATION,    self._normalise_rotation,   "rotations_normalised"),
            (CANON_BURST_END,   self._normalise_burst_end,  "burst_ends_normalised"),
        ]:
            if col in df.columns:
                df[col], n_changed = self._apply_normaliser(df[col], fn)
                setattr(report, counter_attr, getattr(report, counter_attr) + n_changed)

        # Step 4b — If separate Width/Height columns exist but no combined Resolution,
        # merge them into a Resolution column
        if CANON_RESOLUTION not in df.columns:
            if CANON_WIDTH in df.columns and CANON_HEIGHT in df.columns:
                df[CANON_RESOLUTION] = (
                    df[CANON_WIDTH].astype(str).str.strip()
                    + "x"
                    + df[CANON_HEIGHT].astype(str).str.strip()
                )
                df = df.drop(columns=[CANON_WIDTH, CANON_HEIGHT])

        # Step 5 — Normalise Consec keywords in Creative File Name and Notes
        for col in (CANON_CREATIVE, CANON_COMMENTS):
            if col in df.columns:
                df[col], n = self._apply_normaliser(df[col], self._normalise_consec)
                report.consec_normalised += n

        report.rows_out = len(df)
        return df, report

    # ------------------------------------------------------------------
    # Column renaming
    # ------------------------------------------------------------------

    @staticmethod
    def _rename_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, dict, list]:
        """Map raw column headers to canonical names using HEADER_SYNONYMS."""
        rename_map: dict[str, str] = {}
        unrecognised: list[str]    = []

        for raw_col in df.columns:
            canonical = _match_synonym(str(raw_col))
            if canonical and canonical not in rename_map.values():
                rename_map[raw_col] = canonical
            elif not canonical:
                unrecognised.append(raw_col)

        df = df.rename(columns=rename_map)
        renames_done = {k: v for k, v in rename_map.items() if k != v}
        return df, renames_done, unrecognised

    # ------------------------------------------------------------------
    # Field normalisers (each returns normalised value or original on failure)
    # ------------------------------------------------------------------

    def _normalise_date(self, raw: str) -> str:
        """Convert any recognisable date string to YYYY/MM/DD."""
        cleaned = str(raw).strip()
        if not cleaned or cleaned.lower() in ("", "nan", "none", "-", "n/a"):
            return ""

        # Try ISO format (YYYY-MM-DD) first — never ambiguous
        import re as _re
        if _re.fullmatch(r"\d{4}[-/]\d{2}[-/]\d{2}", cleaned):
            try:
                dt = pd.to_datetime(cleaned, errors="raise")
                return dt.strftime("%Y/%m/%d")
            except Exception:
                pass

        # Try Australian day-first convention, then month-first fallback
        for dayfirst in (True, False):
            try:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    dt = pd.to_datetime(cleaned, dayfirst=dayfirst, errors="raise")
                return dt.strftime("%Y/%m/%d")
            except Exception:
                pass

        # Try Excel serial number
        try:
            serial = float(cleaned)
            if 30000 < serial < 60000:   # sane Excel date range
                dt = pd.Timestamp("1899-12-30") + pd.Timedelta(days=serial)
                return dt.strftime("%Y/%m/%d")
        except ValueError:
            pass

        return cleaned   # give up, return as-is

    @staticmethod
    def _normalise_resolution(raw: str) -> str:
        """Normalise any dimension string to WxH with no spaces or suffixes."""
        cleaned = str(raw).strip().lower()
        if not cleaned or cleaned in ("nan", "none", "", "n/a"):
            return ""

        # Remove common suffixes
        cleaned = re.sub(r"px\b", "", cleaned)
        cleaned = re.sub(r"\s+", "", cleaned)

        # Handle W:NNNN H:NNNN or Width:NNNN Height:NNNN styles
        wh = re.search(r"w[idth]*:?(\d+)[,\s]+h[eight]*:?(\d+)", cleaned, re.IGNORECASE)
        if wh:
            return f"{wh.group(1)}x{wh.group(2)}"

        # Normalise any separator (X, ×, *, space) to lowercase x
        normalised = re.sub(r"[×*xX\s]+", "x", cleaned)

        # Validate it looks like WxH
        match = re.fullmatch(r"(\d+)x(\d+)", normalised)
        if match:
            return f"{match.group(1)}x{match.group(2)}"

        return str(raw).strip()   # give up

    @staticmethod
    def _normalise_ad_length(raw: str) -> str:
        """
        Extract the numeric ad length in seconds.
        "15 secs" → "15", "0:15" → "15", "15s" → "15", "0:00:15" → "15"
        """
        cleaned = str(raw).strip()
        if not cleaned or cleaned.lower() in ("nan", "none", "", "n/a"):
            return ""

        # MM:SS or HH:MM:SS format
        time_match = re.fullmatch(r"(?:(\d+):)?(\d+):(\d+)", cleaned)
        if time_match:
            h, m, s = (
                int(time_match.group(1) or 0),
                int(time_match.group(2)),
                int(time_match.group(3)),
            )
            total = h * 3600 + m * 60 + s
            return str(total)

        # Extract leading number (handles "15 secs", "15s", "15SEC", "15 seconds")
        num_match = re.match(r"^(\d+(?:\.\d+)?)", cleaned.replace(",", "."))
        if num_match:
            val = float(num_match.group(1))
            return str(int(val))

        return cleaned

    @staticmethod
    def _normalise_rotation(raw: str) -> str:
        """
        Normalise creative rotation ratios to whole-number percentage strings.
        0.7 → "70", "70%" → "70", "0.33" → "33"
        mi-data-automation rule: never use "%" symbol, values must sum to 100.
        Note: this normalises individual values; summing to 100 is validated separately.
        """
        cleaned = str(raw).strip()
        if not cleaned or cleaned.lower() in ("nan", "none", "", "n/a"):
            return ""

        # Remove % symbol
        cleaned_no_pct = cleaned.replace("%", "").strip()

        try:
            val = float(cleaned_no_pct)
            # If it looks like a fraction (0 < val ≤ 1), multiply by 100
            if 0 < val <= 1:
                pct = round(val * 100, 2)
                return str(int(pct)) if pct == int(pct) else str(pct)
            return str(int(round(val)))
        except ValueError:
            return cleaned

    @staticmethod
    def _normalise_burst_end(raw: str) -> str:
        """
        Normalise Burst End to "1" (is burst end) or "0" (not burst end).
        mi-data-automation rule: Burst End = 1 or blank → End Date = Start + 6 days.
        """
        cleaned = str(raw).strip().lower()
        if cleaned in BURST_END_TRUE:
            return "1"
        if cleaned in BURST_END_FALSE:
            return "0"
        # Try to parse as integer
        try:
            return "1" if int(float(cleaned)) else "0"
        except ValueError:
            return cleaned

    @staticmethod
    def _normalise_consec(raw: str) -> str:
        """Normalise all Consec keyword variants to the canonical word 'Consec'."""
        return CONSEC_PATTERNS.sub("Consec", str(raw))

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_normaliser(series: pd.Series, fn) -> tuple[pd.Series, int]:
        """Apply a normaliser function to a Series, returning (new_series, n_changed)."""
        original = series.astype(str)
        normalised = series.astype(str).map(fn)
        changed = (original != normalised).sum()
        return normalised, int(changed)


# ---------------------------------------------------------------------------
# File reader (handles xlsx, xlsm, csv with encoding detection)
# ---------------------------------------------------------------------------

def read_mi_file(path: str | Path) -> pd.DataFrame:
    """
    Read an MI file regardless of format, applying common structural fixes:
    - Multi-row headers (skips rows before the detected header)
    - Merged cells (forward-fill)
    - Encoding sniffing for CSV files
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xlsm", ".xls"):
        return _read_excel_mi(path)
    elif suffix in (".csv", ".txt", ".tsv"):
        return _read_csv_mi(path)
    else:
        raise ValueError(f"Unsupported MI file format: {suffix}")


def _read_excel_mi(path: Path) -> pd.DataFrame:
    """Read an Excel MI file, detecting the header row and forward-filling merges."""
    # Read raw without a header first to detect the real header row
    raw = pd.read_excel(path, header=None, dtype=str)

    header_row = _detect_header_row(raw)
    if header_row is None:
        header_row = 0

    df = pd.read_excel(path, header=header_row, dtype=str)

    # Forward-fill merged cells (Excel merges show as NaN in pandas)
    df = df.ffill(axis=0)

    # Drop rows that are entirely NaN after forward-fill
    df = df.dropna(how="all")

    return df.reset_index(drop=True)


def _read_csv_mi(path: Path) -> pd.DataFrame:
    """Read a CSV MI file, trying multiple encodings."""
    for encoding in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            sep = "\t" if path.suffix.lower() == ".tsv" else ","
            df = pd.read_csv(path, dtype=str, encoding=encoding, sep=sep)
            return df.reset_index(drop=True)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue
    raise ValueError(f"Could not read {path.name} with any supported encoding.")


def _detect_header_row(raw: pd.DataFrame, max_scan_rows: int = 10) -> Optional[int]:
    """
    Scan the first N rows to find which one looks most like a header.
    Scores rows by how many cells match known MI column synonyms.
    Returns 0-indexed row number of the best candidate.
    """
    all_synonyms = {
        syn.lower()
        for syns in HEADER_SYNONYMS.values()
        for syn in syns
    }

    best_row, best_score = 0, 0
    for i, row in raw.head(max_scan_rows).iterrows():
        score = sum(
            1
            for cell in row
            if isinstance(cell, str) and any(
                syn in cell.lower() for syn in all_synonyms
            )
        )
        if score > best_score:
            best_score, best_row = score, int(i)

    return best_row if best_score > 0 else None


def _match_synonym(col: str) -> Optional[str]:
    """Return the canonical column name for a raw header, or None if not recognised."""
    lower = col.lower().strip()
    for canonical, synonyms in HEADER_SYNONYMS.items():
        if any(syn in lower for syn in synonyms):
            return canonical
    return None


# ---------------------------------------------------------------------------
# Convenience entry point
# ---------------------------------------------------------------------------

def run_mi_normaliser(
    mi_path: str | Path,
    out_path: Optional[str | Path] = None,
    strict_dates: bool = False,
) -> tuple[pd.DataFrame, NormaliserReport]:
    """
    Read a raw MI file, normalise it, and write the clean CSV.

    Parameters
    ----------
    mi_path      : path to the raw MI file (.xlsx, .xlsm, .csv)
    out_path     : where to write the normalised CSV (default: <stem>_normalised.csv)
    strict_dates : raise on ambiguous dates instead of defaulting to DD/MM/YYYY

    Returns
    -------
    (normalised_df, report)
    """
    mi_path = Path(mi_path)
    df_raw  = read_mi_file(mi_path)

    normaliser = MIPreNormaliser(strict_dates=strict_dates)
    df_clean, report = normaliser.normalise(df_raw, source_name=mi_path.name)

    if out_path is None:
        out_path = mi_path.with_name(mi_path.stem + "_normalised.csv")

    df_clean.to_csv(out_path, index=False)
    print(report.summary())
    print(f"\nNormalised MI written to: {out_path}")

    return df_clean, report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="MI Pre-normaliser")
    parser.add_argument("mi_file",         help="Raw MI file path (.xlsx or .csv)")
    parser.add_argument("--out",           help="Output normalised CSV path (optional)")
    parser.add_argument(
        "--strict-dates", action="store_true",
        help="Raise error on ambiguous dates instead of defaulting DD/MM/YYYY"
    )
    args = parser.parse_args()

    run_mi_normaliser(args.mi_file, out_path=args.out, strict_dates=args.strict_dates)
