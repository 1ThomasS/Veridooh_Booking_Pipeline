"""
phase3_format_resolver.py
--------------------------
Phase 3 of the Veridooh BFK pipeline.

Sits between Phase 1+2 (io_to_bkf.py) and Phase 4 (mi-data-automation).

Responsibilities
----------------
1. Detect panels with missing or unknown Format values.
2. Auto-query Metabase to resolve those formats.
3. Validate and correct Screen size/resolution to strict WxH order.
4. Emit a resolution report so the booker can review every change.

Usage
-----
    from phase3_format_resolver import FormatResolver, run_phase3

    # With Metabase (full mode)
    from utils.metabase_client import MetabaseClient
    client  = MetabaseClient()
    bkf_out, report = run_phase3("path/to/bkf_draft.csv", metabase=client)

    # Dry-run / offline (flags panels but does not resolve)
    bkf_out, report = run_phase3("path/to/bkf_draft.csv", metabase=None)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# Column names — must match the exact BKF template headers from bkf-workflow
# ---------------------------------------------------------------------------

COL_PANEL_ID   = "Panel ID (or Player ID or Site ID)"
COL_FORMAT     = "Format"
COL_SCREEN_SZ  = "Screen size/ resolution"
COL_CREATIVE   = "Creative File Name"

# Tokens that indicate the Format cell is effectively empty
EMPTY_FORMAT_TOKENS = {"", "n/a", "tbc", "unknown", "nan", "none", "-"}

# Known format strings that the Format Checker accepts as valid.
# TODO: expand this list as you encounter more valid formats on the job.
# This acts as a quick local check before hitting Metabase.
KNOWN_VALID_FORMATS: set[str] = {
    "classic portrait",
    "classic landscape",
    "full motion portrait",
    "full motion landscape",
    "super portrait",
    "billboard",
    "spectacular",
    "trivision",
    "classic",
}


# ---------------------------------------------------------------------------
# Resolution report
# ---------------------------------------------------------------------------

@dataclass
class PanelResolution:
    panel_id: str
    original_format: str
    resolved_format: Optional[str]
    original_dim: str
    resolved_dim: Optional[str]
    dim_swapped: bool          = False
    from_metabase: bool        = False
    needs_manual_review: bool  = False
    note: str                  = ""


@dataclass
class ResolutionReport:
    total_panels: int              = 0
    already_valid: int             = 0
    resolved_via_metabase: int     = 0
    dim_corrections: int           = 0
    still_unresolved: int          = 0
    needs_manual_review: int       = 0
    items: list[PanelResolution]   = field(default_factory=list)

    def summary(self) -> str:
        lines = [
            "=" * 60,
            "Phase 3 — Format Resolver Report",
            "=" * 60,
            f"  Total panels          : {self.total_panels}",
            f"  Already valid         : {self.already_valid}",
            f"  Resolved via Metabase : {self.resolved_via_metabase}",
            f"  Dimension corrections : {self.dim_corrections}",
            f"  Still unresolved      : {self.still_unresolved}",
            f"  Needs manual review   : {self.needs_manual_review}",
            "=" * 60,
        ]
        if self.still_unresolved > 0 or self.needs_manual_review > 0:
            lines.append("\nAction required:")
            for item in self.items:
                if item.needs_manual_review or item.resolved_format is None:
                    status = "UNRESOLVED" if item.resolved_format is None else "REVIEW"
                    lines.append(
                        f"  [{status}] Panel {item.panel_id!r}: "
                        f"format={item.original_format!r} | {item.note}"
                    )
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Core resolver
# ---------------------------------------------------------------------------

class FormatResolver:
    """
    Validates and repairs Format and Screen size/resolution columns in a BKF dataframe.

    Parameters
    ----------
    metabase : MetabaseClient or None
        Pass a configured MetabaseClient to enable auto-resolution.
        Pass None to run in flag-only mode (useful for dry runs or testing).
    """

    def __init__(self, metabase=None):
        self.metabase = metabase

    def resolve(self, df: pd.DataFrame) -> tuple[pd.DataFrame, ResolutionReport]:
        """
        Process the BKF dataframe in-place (operates on a copy).

        Returns
        -------
        df     : pd.DataFrame  Updated BKF with Format and Screen size filled in.
        report : ResolutionReport  Full audit trail of every change.
        """
        df = df.copy()
        report = ResolutionReport(total_panels=len(df))

        # Prefetch all panel IDs from Metabase in one shot (avoids N round trips)
        if self.metabase is not None:
            panel_ids = df[COL_PANEL_ID].dropna().astype(str).tolist()
            self.metabase.prefetch(panel_ids)

        for idx, row in df.iterrows():
            panel_id    = str(row.get(COL_PANEL_ID, "")).strip()
            raw_format  = str(row.get(COL_FORMAT,    "")).strip()
            raw_dim     = str(row.get(COL_SCREEN_SZ, "")).strip()

            item = PanelResolution(
                panel_id        = panel_id,
                original_format = raw_format,
                resolved_format = raw_format if raw_format else None,
                original_dim    = raw_dim,
                resolved_dim    = raw_dim if raw_dim else None,
            )

            # 1. Validate/resolve Format
            if self._is_empty(raw_format):
                item = self._resolve_format(item, df, idx)
            elif not self._is_known_valid(raw_format):
                # Format exists but isn't in our known-valid list — flag for review
                item.needs_manual_review = True
                item.note = f"Format {raw_format!r} not in known-valid list; verify manually"
                report.needs_manual_review += 1
            else:
                report.already_valid += 1

            # 2. Validate dimension order (must be WxH)
            corrected_dim, was_swapped = self._validate_dimension(
                item.resolved_dim or raw_dim, item.resolved_format
            )
            if was_swapped:
                item.resolved_dim = corrected_dim
                item.dim_swapped  = True
                item.note += f" | Dimension auto-corrected: {raw_dim} → {corrected_dim}"
                report.dim_corrections += 1

            # 3. Write resolved values back to the dataframe
            if item.resolved_format:
                df.at[idx, COL_FORMAT] = item.resolved_format
            if item.resolved_dim:
                df.at[idx, COL_SCREEN_SZ] = item.resolved_dim

            report.items.append(item)

        # Tally unresolved
        for item in report.items:
            if item.resolved_format is None:
                report.still_unresolved += 1

        return df, report

    # ------------------------------------------------------------------
    # Format resolution helpers
    # ------------------------------------------------------------------

    def _resolve_format(
        self, item: PanelResolution, df: pd.DataFrame, idx: int
    ) -> PanelResolution:
        """Attempt to fill a missing Format from Metabase, then flag if still unknown."""
        if self.metabase is not None:
            mb_format = self.metabase.get_panel_format(item.panel_id)
            if mb_format:
                item.resolved_format = mb_format.strip()
                item.from_metabase   = True
                item.note = f"Format resolved from Metabase: {mb_format!r}"
                return item

        # Metabase didn't help — flag for manual Metabase lookup
        item.needs_manual_review = True
        item.note = (
            "Format missing and could not be resolved. "
            "Look up Panel ID in Metabase Master Site List."
        )
        return item

    # ------------------------------------------------------------------
    # Dimension validation helpers
    # ------------------------------------------------------------------

    def _validate_dimension(
        self, raw: str, fmt: Optional[str]
    ) -> tuple[str, bool]:
        """
        Ensure the dimension string is in WxH (width × height) order.

        Strategy:
        - Parse the raw string into two integers.
        - Use the resolved Format (if available) to infer portrait vs landscape.
        - Swap if the current order appears to be HxW.

        Returns (corrected_dim, was_swapped).
        """
        if not raw or raw.lower() in ("", "nan", "n/a"):
            return raw, False

        w, h = _parse_dim(raw)
        if w is None or h is None:
            return raw, False   # Can't parse — leave as-is, booker reviews

        # Determine expected orientation from Format string
        expected_portrait = _format_suggests_portrait(fmt)

        # Current values: first number = assumed Width, second = assumed Height
        current_is_portrait = w < h  # e.g. 1080x1920 → portrait ✓

        if expected_portrait is None:
            # No orientation hint — can't determine correct order
            return f"{w}x{h}", False

        if expected_portrait and not current_is_portrait:
            # Format says portrait but numbers look landscape — swap
            return f"{h}x{w}", True

        if not expected_portrait and current_is_portrait:
            # Format says landscape but numbers look portrait — swap
            return f"{h}x{w}", True

        # Order looks correct
        return f"{w}x{h}", False

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    @staticmethod
    def _is_empty(value: str) -> bool:
        return value.lower().strip() in EMPTY_FORMAT_TOKENS

    @staticmethod
    def _is_known_valid(value: str) -> bool:
        return value.lower().strip() in KNOWN_VALID_FORMATS


# ---------------------------------------------------------------------------
# Dimension parsing helpers (module-level)
# ---------------------------------------------------------------------------

_DIM_RE = re.compile(r"(\d+)\s*[xX×]\s*(\d+)")


def _parse_dim(raw: str) -> tuple[Optional[int], Optional[int]]:
    """Extract (width, height) integers from a raw dimension string, or (None, None)."""
    match = _DIM_RE.search(raw.strip())
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None


def _format_suggests_portrait(fmt: Optional[str]) -> Optional[bool]:
    """
    Return True if the format name implies a portrait screen,
    False if landscape, None if unknown.
    """
    if not fmt:
        return None
    lower = fmt.lower()
    if "portrait" in lower:
        return True
    if "landscape" in lower or "billboard" in lower or "spectacular" in lower:
        return False
    return None     # e.g. "Classic" — could be either


# ---------------------------------------------------------------------------
# Convenience entry point
# ---------------------------------------------------------------------------

def run_phase3(
    bkf_csv_path: str | Path,
    metabase=None,
    out_csv_path: Optional[str | Path] = None,
) -> tuple[pd.DataFrame, ResolutionReport]:
    """
    Load a Phase 1+2 BKF CSV, resolve formats, and write the updated CSV.

    Parameters
    ----------
    bkf_csv_path : path to the io_to_bkf.py output CSV
    metabase     : configured MetabaseClient, or None for dry-run
    out_csv_path : where to write the resolved BKF (default: <input>_phase3.csv)

    Returns
    -------
    (resolved_df, report)
    """
    bkf_path = Path(bkf_csv_path)
    df = pd.read_csv(bkf_path, dtype=str).fillna("")

    resolver   = FormatResolver(metabase=metabase)
    df_out, report = resolver.resolve(df)

    if out_csv_path is None:
        out_csv_path = bkf_path.with_name(bkf_path.stem + "_phase3.csv")

    df_out.to_csv(out_csv_path, index=False)
    print(report.summary())
    print(f"\nResolved BKF written to: {out_csv_path}")

    return df_out, report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import argparse

    parser = argparse.ArgumentParser(description="Phase 3 Format Resolver")
    parser.add_argument("bkf_csv", help="Path to Phase 1+2 BKF CSV")
    parser.add_argument("--out",   help="Output CSV path (optional)")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Flag issues without querying Metabase"
    )
    args = parser.parse_args()

    client = None
    if not args.dry_run:
        try:
            from utils.metabase_client import MetabaseClient
            client = MetabaseClient()
        except Exception as e:
            print(f"[Warning] Could not init Metabase client: {e}")
            print("[Warning] Running in flag-only mode.")

    run_phase3(args.bkf_csv, metabase=client, out_csv_path=args.out)
