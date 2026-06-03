"""
phase45_merger.py
------------------
Phase 4.5 of the Veridooh BFK pipeline.

Sits between Phase 4 (mi-data-automation output) and Phase 5 (Post-processor).

Responsibilities
----------------
1. Merge the Phase 4 creative allocation output into the Phase 3 resolved BKF.
2. Apply do-not-book filtering (remove or flag panels that must not be booked).
3. Handle partial fill: where the MI only covers certain creative dimensions,
   panels with no matching creative are left blank (already handled by Phase 4,
   confirmed and audited here).
4. Emit a merge report for the booker.

Usage
-----
    from phase45_merger import run_phase45

    merged_df, report = run_phase45(
        bkf_phase3_csv  = "campaign_phase3.csv",
        creative_csv    = "Output.csv",         # Phase 4 mi-data-automation output
        do_not_book_ids = ["P001", "P045"],     # optional explicit exclusion list
        out_csv_path    = "campaign_phase45.csv",
    )
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# BKF column names (must match bkf-workflow SKILL.md)
# ---------------------------------------------------------------------------

BKF_PANEL_ID      = "Panel ID (or Player ID or Site ID)"
BKF_CREATIVE      = "Creative File Name"
BKF_PLAY_INSTR    = "Play Instructions"
BKF_START_DATE    = "Creative Start Date"
BKF_END_DATE      = "Creative End Date"
BKF_COMMENTS      = "Other Comments"

# Phase 4 (mi-data-automation) output column names (must match SKILL.md)
MI_PANEL_ID       = "Panel ID"
MI_CREATIVE       = "Creative File Name"
MI_PLAY_INSTR     = "Play Instructions"
MI_START_DATE     = "Start Date (MI)"
MI_END_DATE       = "End Date (MI)"
MI_BKF_START      = "Booking Start Date (BKF)"
MI_BKF_END        = "Booking End Date (BKF)"
MI_COMMENTS       = "Comments"

# Tokens that indicate a creative slot was intentionally left blank
BLANK_CREATIVE_TOKENS = {"", "nan", "none", "n/a"}

# Column used to flag do-not-book panels in the IO/BKF.
# If your IO marks these with a column value, set this.
# If you prefer to pass an explicit list of Panel IDs, set this to None.
# TODO: confirm the exact column name and value with your team.
DNB_FLAG_COLUMN = None        # e.g. "Booking Type"
DNB_FLAG_VALUE  = None        # e.g. "Do Not Book"


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

@dataclass
class MergeReport:
    total_bkf_rows: int      = 0
    creatives_filled: int    = 0
    blank_partial_fill: int  = 0
    do_not_book_removed: int = 0
    unmatched_panels: int    = 0
    notes: list[str]         = field(default_factory=list)

    def summary(self) -> str:
        lines = [
            "=" * 60,
            "Phase 4.5 — BFK Merger Report",
            "=" * 60,
            f"  BKF rows in             : {self.total_bkf_rows}",
            f"  Creatives filled        : {self.creatives_filled}",
            f"  Blank (partial fill)    : {self.blank_partial_fill}",
            f"  Do-not-book removed     : {self.do_not_book_removed}",
            f"  Unmatched panels        : {self.unmatched_panels}",
            "=" * 60,
        ]
        if self.notes:
            lines.append("\nNotes:")
            for note in self.notes:
                lines.append(f"  · {note}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Core merger
# ---------------------------------------------------------------------------

class BFKMerger:
    """
    Merge Phase 3 BKF with Phase 4 creative allocation output.

    Parameters
    ----------
    do_not_book_ids : list of Panel IDs to exclude from the final BKF.
                      Takes precedence over DNB_FLAG_COLUMN matching.
    prefer_mi_dates : if True, replace BKF Creative Start/End dates with MI dates
                      where the MI provides them (recommended — MI dates are more
                      authoritative for the actual play window).
    """

    def __init__(
        self,
        do_not_book_ids: Optional[list[str]] = None,
        prefer_mi_dates: bool = True,
    ):
        self.do_not_book_ids = {
            str(pid).strip() for pid in (do_not_book_ids or [])
        }
        self.prefer_mi_dates = prefer_mi_dates

    def merge(
        self,
        bkf_df: pd.DataFrame,
        creative_df: pd.DataFrame,
    ) -> tuple[pd.DataFrame, MergeReport]:
        """
        Merge creative allocation into the BKF.

        Parameters
        ----------
        bkf_df      : Phase 3 resolved BKF dataframe
        creative_df : Phase 4 mi-data-automation Output.csv dataframe

        Returns
        -------
        (merged_df, report)
        """
        bkf  = bkf_df.copy()
        crea = creative_df.copy()
        report = MergeReport(total_bkf_rows=len(bkf))

        # Normalise Panel ID columns for matching
        bkf[BKF_PANEL_ID]  = bkf[BKF_PANEL_ID].astype(str).str.strip()
        crea[MI_PANEL_ID]   = crea[MI_PANEL_ID].astype(str).str.strip()

        # Build a lookup dict: panel_id → creative row (Phase 4 output)
        # Phase 4 guarantees same row count as BKF, but we match by Panel ID
        # in case row order drifts.
        creative_lookup: dict[str, pd.Series] = {}
        for _, row in crea.iterrows():
            pid = row[MI_PANEL_ID]
            if pid in creative_lookup:
                report.notes.append(
                    f"Duplicate Panel ID {pid!r} in Phase 4 output — "
                    "using first occurrence"
                )
            else:
                creative_lookup[pid] = row

        # Apply do-not-book filter from flag column (if configured)
        if DNB_FLAG_COLUMN and DNB_FLAG_COLUMN in bkf.columns:
            dnb_from_column = set(
                bkf.loc[
                    bkf[DNB_FLAG_COLUMN].str.strip().str.lower()
                    == str(DNB_FLAG_VALUE).lower(),
                    BKF_PANEL_ID,
                ].tolist()
            )
            self.do_not_book_ids.update(dnb_from_column)
            if dnb_from_column:
                report.notes.append(
                    f"Do-not-book panels from column {DNB_FLAG_COLUMN!r}: "
                    f"{sorted(dnb_from_column)}"
                )

        # Remove do-not-book rows
        if self.do_not_book_ids:
            before = len(bkf)
            bkf = bkf[~bkf[BKF_PANEL_ID].isin(self.do_not_book_ids)]
            removed = before - len(bkf)
            report.do_not_book_removed = removed
            if removed:
                report.notes.append(
                    f"Removed {removed} do-not-book panel(s): "
                    f"{sorted(self.do_not_book_ids)}"
                )

        # Merge creatives row by row
        rows_out = []
        for _, bkf_row in bkf.iterrows():
            pid        = bkf_row[BKF_PANEL_ID]
            crea_row   = creative_lookup.get(pid)

            merged_row = bkf_row.copy()

            if crea_row is None:
                report.unmatched_panels += 1
                report.notes.append(
                    f"Panel {pid!r} has no Phase 4 creative row — "
                    "Creative File Name left blank"
                )
            else:
                # Fill Creative File Name + Play Instructions from Phase 4
                creative_val = str(crea_row.get(MI_CREATIVE, "")).strip()
                play_val     = str(crea_row.get(MI_PLAY_INSTR, "")).strip()
                mi_comment   = str(crea_row.get(MI_COMMENTS, "")).strip()

                merged_row[BKF_CREATIVE]   = creative_val
                merged_row[BKF_PLAY_INSTR] = play_val

                # Append MI comment to Other Comments (preserve existing BKF notes)
                existing_comment = str(merged_row.get(BKF_COMMENTS, "")).strip()
                if mi_comment and mi_comment.lower() not in ("", "nan"):
                    merged_row[BKF_COMMENTS] = (
                        f"{existing_comment} | {mi_comment}".lstrip(" |")
                        if existing_comment
                        else mi_comment
                    )

                # Prefer MI dates where available (more authoritative)
                if self.prefer_mi_dates:
                    mi_start = str(crea_row.get(MI_START_DATE, "")).strip()
                    mi_end   = str(crea_row.get(MI_END_DATE, "")).strip()
                    if mi_start and mi_start.lower() not in ("", "nan"):
                        merged_row[BKF_START_DATE] = mi_start
                    if mi_end and mi_end.lower() not in ("", "nan"):
                        merged_row[BKF_END_DATE] = mi_end

                # Count fill status
                if creative_val.lower() in BLANK_CREATIVE_TOKENS:
                    report.blank_partial_fill += 1
                else:
                    report.creatives_filled += 1

            rows_out.append(merged_row)

        merged_df = pd.DataFrame(rows_out, columns=bkf.columns)
        return merged_df, report


# ---------------------------------------------------------------------------
# Convenience entry point
# ---------------------------------------------------------------------------

def run_phase45(
    bkf_phase3_csv: str | Path,
    creative_csv: str | Path,
    do_not_book_ids: Optional[list[str]] = None,
    out_csv_path: Optional[str | Path] = None,
    prefer_mi_dates: bool = True,
) -> tuple[pd.DataFrame, MergeReport]:
    """
    Load Phase 3 BKF and Phase 4 creative output, merge, and write the result.

    Parameters
    ----------
    bkf_phase3_csv   : path to phase3 resolved BKF CSV
    creative_csv     : path to mi-data-automation Output.csv
    do_not_book_ids  : list of Panel IDs to exclude (optional)
    out_csv_path     : output path (default: <bkf_stem>_phase45.csv)
    prefer_mi_dates  : use MI dates over BKF booking dates where available
    """
    bkf_path  = Path(bkf_phase3_csv)
    crea_path = Path(creative_csv)

    bkf_df   = pd.read_csv(bkf_path,  dtype=str).fillna("")
    crea_df  = pd.read_csv(crea_path, dtype=str).fillna("")

    merger = BFKMerger(
        do_not_book_ids=do_not_book_ids,
        prefer_mi_dates=prefer_mi_dates,
    )
    merged_df, report = merger.merge(bkf_df, crea_df)

    if out_csv_path is None:
        out_csv_path = bkf_path.with_name(bkf_path.stem.replace("_phase3", "") + "_phase45.csv")

    merged_df.to_csv(out_csv_path, index=False)
    print(report.summary())
    print(f"\nMerged BKF written to: {out_csv_path}")

    return merged_df, report


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Phase 4.5 BFK Merger")
    parser.add_argument("bkf_csv",       help="Phase 3 BKF CSV")
    parser.add_argument("creative_csv",  help="Phase 4 mi-data-automation Output.csv")
    parser.add_argument("--out",         help="Output CSV path")
    parser.add_argument(
        "--do-not-book", nargs="*", default=[],
        metavar="PANEL_ID",
        help="Panel IDs to exclude (space-separated)"
    )
    parser.add_argument(
        "--keep-bkf-dates", action="store_true",
        help="Use BKF booking dates instead of MI dates"
    )
    args = parser.parse_args()

    run_phase45(
        bkf_phase3_csv  = args.bkf_csv,
        creative_csv    = args.creative_csv,
        do_not_book_ids = args.do_not_book or None,
        out_csv_path    = args.out,
        prefer_mi_dates = not args.keep_bkf_dates,
    )
