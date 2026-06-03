"""
orchestrator.py
---------------
Single entry point for the Veridooh BFK Pipeline Orchestrator.

Chains all phases in order:
    Phase 1+2  (io_to_bkf.py — existing skill, invoked as subprocess)
    Phase 3    format_resolver   ← new
    Phase 4    mi-data-automation (existing skill — must be run separately via Claude)
    Phase 4.5  bfk_merger        ← new
    Phase 5    postprocessor      ← new
    Phase 6    upload_notify      ← stub (add details later)

Usage
-----
    # Full pipeline (requires Phase 4 output to already exist)
    python orchestrator.py \\
        --io        "JCD_Lifeblood_IO.xlsx" \\
        --creative  "Output.csv" \\
        --campaign  "JCD_Lifeblood" \\
        --do-not-book P001 P045

    # Skip Phase 1+2 if you already have a BKF draft
    python orchestrator.py \\
        --bkf-draft "JCD_Lifeblood_BKF.csv" \\
        --creative  "Output.csv" \\
        --campaign  "JCD_Lifeblood"

    # Dry-run (no Metabase queries, flag issues only)
    python orchestrator.py --io "..." --creative "..." --dry-run
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

import pandas as pd

from mi_pre_normaliser       import run_mi_normaliser
from phase3_format_resolver  import run_phase3
from phase45_merger          import run_phase45
from phase5_postprocessor    import run_phase5


# ---------------------------------------------------------------------------
# Pipeline runner
# ---------------------------------------------------------------------------

class Pipeline:
    """
    Orchestrates all phases for a single campaign.

    Parameters
    ----------
    campaign_name   : short name used for output file naming
    work_dir        : directory where all intermediate files are written
    dry_run         : if True, skip Metabase queries (flag-only mode)
    do_not_book_ids : list of Panel IDs to exclude from the final BKF
    prefer_mi_dates : if True, use MI dates over BKF booking dates
    """

    def __init__(
        self,
        campaign_name: str,
        work_dir: str | Path = ".",
        dry_run: bool = False,
        do_not_book_ids: Optional[list[str]] = None,
        prefer_mi_dates: bool = True,
    ):
        self.name           = campaign_name
        self.work_dir       = Path(work_dir)
        self.dry_run        = dry_run
        self.dnb_ids        = do_not_book_ids or []
        self.prefer_mi_dates = prefer_mi_dates
        self.work_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Phase 1+2 — invoke io_to_bkf.py as a subprocess
    # ------------------------------------------------------------------

    def run_phase12(self, io_path: str | Path) -> Path:
        """
        Run the existing io_to_bkf.py script and return the path to the
        generated BKF CSV.

        Assumes io_to_bkf.py is importable from scripts/io_to_bkf.py
        (as per the bkf-workflow skill).
        """
        io_path  = Path(io_path)
        out_path = self.work_dir / f"{self.name}_BKF.csv"

        # Locate the script relative to this file (adjust path if needed)
        script = Path(__file__).parent / "scripts" / "io_to_bkf.py"
        if not script.exists():
            raise FileNotFoundError(
                f"io_to_bkf.py not found at {script}. "
                "Either place it under scripts/ or point to your copy."
            )

        cmd = [
            sys.executable, str(script),
            str(io_path),
            "--out", str(out_path),
        ]

        print(f"\n[Phase 1+2] Running io_to_bkf.py on {io_path.name} …")
        t0 = time.time()
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(result.stderr)
            raise RuntimeError("Phase 1+2 failed — see output above.")

        print(result.stdout)
        print(f"[Phase 1+2] ✓ Done in {time.time()-t0:.1f}s → {out_path.name}")
        return out_path

    # ------------------------------------------------------------------
    # Phase 3.5 — MI pre-normaliser (runs before Phase 4)
    # ------------------------------------------------------------------

    def run_phase35(self, mi_path) -> Path:
        """
        Normalise the raw MI file before it goes into mi-data-automation.
        Fixes column names, date formats, resolution strings, ad lengths,
        rotation ratios, burst end values, and Consec keyword variants.
        """
        out_path = self.work_dir / f"{self.name}_MI_normalised.csv"
        print(f"[Phase 3.5] Normalising MI file: {Path(mi_path).name} ...")
        t0 = time.time()
        _, report = run_mi_normaliser(mi_path, out_path=out_path)
        print(f"[Phase 3.5] Done in {time.time()-t0:.1f}s -> {out_path.name}")
        total_fixes = (
            report.dates_normalised + report.resolutions_normalised
            + report.ad_lengths_normalised + report.rotations_normalised
            + report.burst_ends_normalised + report.consec_normalised
        )
        if total_fixes > 0:
            print(f"[Phase 3.5] {total_fixes} field(s) normalised. Check report before running mi-data-automation.")
        return out_path

        # ------------------------------------------------------------------
    # Phase 3 — Format resolver
    # ------------------------------------------------------------------

    def run_phase3(self, bkf_draft_path: str | Path) -> Path:
        metabase = None
        if not self.dry_run:
            try:
                from utils.metabase_client import MetabaseClient
                metabase = MetabaseClient()
                print("\n[Phase 3] Metabase client connected ✓")
            except Exception as exc:
                print(f"\n[Phase 3] Warning: Metabase unavailable ({exc})")
                print("[Phase 3] Running in flag-only mode.")

        out_path = self.work_dir / f"{self.name}_phase3.csv"
        print(f"\n[Phase 3] Resolving formats for {Path(bkf_draft_path).name} …")
        t0 = time.time()
        _, report = run_phase3(bkf_draft_path, metabase=metabase, out_csv_path=out_path)
        print(f"[Phase 3] ✓ Done in {time.time()-t0:.1f}s → {out_path.name}")

        if report.still_unresolved > 0:
            print(
                f"[Phase 3] ⚠  {report.still_unresolved} panel(s) still need manual "
                "format lookup in Metabase. Check the report above before continuing."
            )
        return out_path

    # ------------------------------------------------------------------
    # Phase 4 — MI-data-automation (external — Claude-run)
    # ------------------------------------------------------------------

    def phase4_instructions(self, bkf_phase3_path: Path, mi_path: str | Path) -> None:
        """
        Phase 4 is run via the mi-data-automation Claude skill, not as Python.
        This method prints the exact instructions the booker needs.
        """
        print(f"""
[Phase 4] ── MI-data-automation (run via Claude) ──────────────────
This phase is handled by the existing mi-data-automation Claude skill.

Upload these three files to Claude:
  1. BKF (master panel list)  → {bkf_phase3_path}
  2. MI (material instructions) → {mi_path}
  3. Live Creative List         → (paste from Veridooh creative dashboard)

Claude will produce: Output.csv
Save Output.csv to: {self.work_dir / "Output.csv"}
Then re-run this script with --creative {self.work_dir / "Output.csv"}
────────────────────────────────────────────────────────────────────
""")

    # ------------------------------------------------------------------
    # Phase 4.5 — BKF Merger
    # ------------------------------------------------------------------

    def run_phase45(
        self,
        bkf_phase3_path: str | Path,
        creative_csv_path: str | Path,
    ) -> Path:
        out_path = self.work_dir / f"{self.name}_phase45.csv"
        print(f"\n[Phase 4.5] Merging creatives into BKF …")
        t0 = time.time()
        _, report = run_phase45(
            bkf_phase3_csv  = bkf_phase3_path,
            creative_csv    = creative_csv_path,
            do_not_book_ids = self.dnb_ids or None,
            out_csv_path    = out_path,
            prefer_mi_dates = self.prefer_mi_dates,
        )
        print(f"[Phase 4.5] ✓ Done in {time.time()-t0:.1f}s → {out_path.name}")
        return out_path

    # ------------------------------------------------------------------
    # Phase 5 — Post-processor
    # ------------------------------------------------------------------

    def run_phase5(self, merged_csv_path: str | Path) -> Path:
        out_xlsx = self.work_dir / f"{self.name}_FINAL.xlsx"
        print(f"\n[Phase 5] Post-processing …")
        t0 = time.time()
        report = run_phase5(merged_csv=merged_csv_path, out_xlsx=out_xlsx)
        print(f"[Phase 5] ✓ Done in {time.time()-t0:.1f}s → {out_xlsx.name}")

        if report.compliance_errors > 0:
            print(
                f"\n[Phase 5] ⚠  STOP — {report.compliance_errors} compliance error(s). "
                "Do NOT upload until resolved."
            )
        return out_xlsx

    # ------------------------------------------------------------------
    # Phase 6 — Upload & notify (stub)
    # ------------------------------------------------------------------

    def run_phase6(self, final_xlsx_path: Path) -> None:
        """
        TODO: implement once you have:
          - Google Sheets API credentials (service account JSON)
          - Monday.com API token
          - The target Monday.com board ID and column IDs

        For now, prints upload instructions manually.
        """
        print(f"""
[Phase 6] ── Upload & Notify (manual until API is configured) ──────
  1. Upload {final_xlsx_path.name} to the campaign's Google Sheets BKF tab.
  2. Copy the Google Sheets URL.
  3. Paste the URL into the Monday.com item → BKF Link column.
  4. Monday.com auto-bot will run its quality check.
─────────────────────────────────────────────────────────────────────
""")

    # ------------------------------------------------------------------
    # Full pipeline entry point
    # ------------------------------------------------------------------

    def run(
        self,
        io_path: Optional[str | Path] = None,
        bkf_draft_path: Optional[str | Path] = None,
        mi_path: Optional[str | Path] = None,
        creative_csv_path: Optional[str | Path] = None,
        skip_phase12: bool = False,
        skip_phase3: bool = False,
    ) -> Path:
        """
        Run the full pipeline from whatever starting point you have.

        At minimum you need either:
          (a) io_path + creative_csv_path  → full run
          (b) bkf_draft_path + creative_csv_path → skip Phase 1+2

        If creative_csv_path is absent, the pipeline pauses at Phase 4
        and prints instructions for the mi-data-automation step.
        """
        print(f"\n{'='*60}")
        print(f"  Veridooh BFK Pipeline — Campaign: {self.name}")
        print(f"{'='*60}")

        # --- Phase 1+2 ---
        if bkf_draft_path:
            p12_out = Path(bkf_draft_path)
            print(f"\n[Phase 1+2] Using existing BKF draft: {p12_out.name}")
        elif io_path and not skip_phase12:
            p12_out = self.run_phase12(io_path)
        else:
            raise ValueError(
                "Provide either --io (IO file) or --bkf-draft (existing BKF draft)."
            )

        # --- Phase 3 ---
        if skip_phase3:
            p3_out = p12_out
            print("\n[Phase 3] Skipped.")
        else:
            p3_out = self.run_phase3(p12_out)

        # --- Phase 3.5 — MI pre-normaliser (if MI file provided) ---
        normalised_mi_path = None
        if mi_path:
            normalised_mi_path = self.run_phase35(mi_path)

        # --- Phase 4 (external) ---
        if creative_csv_path is None:
            self.phase4_instructions(p3_out, normalised_mi_path or mi_path or "<MI file>")
            print("\nPipeline paused at Phase 4. Re-run with --creative <Output.csv>.")
            return p3_out

        # --- Phase 4.5 ---
        p45_out = self.run_phase45(p3_out, creative_csv_path)

        # --- Phase 5 ---
        final_xlsx = self.run_phase5(p45_out)

        # --- Phase 6 ---
        self.run_phase6(final_xlsx)

        print(f"\n{'='*60}")
        print(f"  Pipeline complete. Final BKF: {final_xlsx}")
        print(f"{'='*60}\n")

        return final_xlsx


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Veridooh BFK Pipeline Orchestrator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--campaign",     required=True, help="Campaign name (used for file naming)")
    parser.add_argument("--io",           help="IO file path (.xlsx or .csv)")
    parser.add_argument("--bkf-draft",   help="Existing Phase 1+2 BKF draft CSV (skip Phase 1+2)")
    parser.add_argument("--mi",           help="MI file path (for Phase 4 instructions)")
    parser.add_argument("--creative",     help="Phase 4 Output.csv from mi-data-automation")
    parser.add_argument("--out-dir",      default=".", help="Output directory")
    parser.add_argument("--do-not-book",  nargs="*", default=[], metavar="PANEL_ID",
                        help="Panel IDs to exclude (space-separated)")
    parser.add_argument("--dry-run",      action="store_true",
                        help="Flag issues without querying Metabase")
    parser.add_argument("--skip-phase3",  action="store_true",
                        help="Skip format resolution (use BKF as-is from Phase 1+2)")
    parser.add_argument("--keep-bkf-dates", action="store_true",
                        help="Use BKF booking dates instead of MI dates")

    args = parser.parse_args()

    pipeline = Pipeline(
        campaign_name   = args.campaign,
        work_dir        = args.out_dir,
        dry_run         = args.dry_run,
        do_not_book_ids = args.do_not_book or None,
        prefer_mi_dates = not args.keep_bkf_dates,
    )

    pipeline.run(
        io_path          = args.io,
        bkf_draft_path   = args.bkf_draft,
        mi_path          = args.mi,
        creative_csv_path = args.creative,
        skip_phase3      = args.skip_phase3,
    )
