# CLAUDE.md

Guidance for working in this repo. Read this before editing pipeline code.

## What this is

An automated pre-ingestion pipeline for Veridooh's OOH booking workflow. It turns a
media-agency **IO** + **MI** + **Live Creative List** into an upload-ready **BKF**
(booking form) Excel file. The final file is uploaded to Google Sheets and linked on
Monday.com, where an auto-bot quality-checks it before booking.

This repo automates the steps a booker does by hand: extract IO details into the BKF
template, normalise the MI, pull creatives + play instructions, resolve panel formats,
apply do-not-book and partial-fill rules, then sort and burst-highlight the result.

## Architecture — phase chain

[orchestrator.py](orchestrator.py) chains everything. Phases:

| Phase | Module | Role | Status |
|------|--------|------|--------|
| 1+2  | `scripts/io_to_bkf.py` | IO → BKF draft | **External** — the `bkf-workflow` Claude skill. Not in this repo; orchestrator shells out to it. |
| 3.5  | [mi_pre_normaliser.py](mi_pre_normaliser.py) | Clean raw agency MI (headers, dates, resolutions, rotations, burst, consec) | In repo |
| 3    | [phase3_format_resolver.py](phase3_format_resolver.py) | Fill missing Format via Metabase; force dimensions to WxH | In repo |
| 4    | mi-data-automation | Creatives + play instructions → `Output.csv` | **External** — run separately via Claude skill. Pipeline pauses here. |
| 4.5  | [phase45_merger.py](phase45_merger.py) | Merge creatives into BKF; do-not-book filter; partial fill | In repo |
| 5    | [phase5_postprocessor.py](phase5_postprocessor.py) | Sort by date; burst-week highlight; compliance; styled `.xlsx` | In repo |
| 6    | `Pipeline.run_phase6()` | Google Sheets + Monday.com upload | **Stub** — prints manual instructions only |

Data flow between phases is **CSV**; only the final Phase 5 output is `.xlsx`. Each phase
also returns a dataclass report object (`*Report`) with a `.summary()` for the booker.

## Environment & running

```bash
# Python 3.10, deps already in .venv (pandas, openpyxl, requests)
.venv/bin/python <script>            # always use the venv interpreter
```

Full run (resume after Phase 4 produced an Output.csv):
```bash
.venv/bin/python orchestrator.py --campaign NAME \
    --bkf-draft draft.csv --creative Output.csv [--do-not-book P1 P2] [--dry-run]
```

Individual phases (each has a CLI):
```bash
.venv/bin/python mi_pre_normaliser.py <MI.xlsx> --out mi_norm.csv
.venv/bin/python phase3_format_resolver.py <bkf.csv> --out p3.csv --dry-run
.venv/bin/python phase45_merger.py <p3.csv> <Output.csv> --do-not-book P001
.venv/bin/python phase5_postprocessor.py <p45.csv> --out FINAL.xlsx
```

`--dry-run` skips Metabase (flag-only). There is no test suite; validate changes by
running phases against real campaign files (see below).

## Test data (not in repo)

Real campaigns live under `~/Documents/Veridooh/Booking_Process/`:
- `IKEA_AU356371/` — has a real `IKEA_WinterSale_BKF.xlsx` (good Phase-3 input).
- `SBS_WC_36579472/` — has a real `SBS_FIFA_WC_MI_Output.csv` (Phase-4 output shape) +
  OASIS IO export (good Phase 4.5 → 5 input).

## CRITICAL: column-name contracts

Phases pass data by **exact column-header strings**, defined as constants at the top of
each module. They must stay in sync across files and match the BKF template / SKILL.md.
Breaking these silently produces blank columns or unmatched rows.

- BKF panel ID header: `"Panel ID (or Player ID or Site ID)"`
- Phase-4 output panel ID header: `"Panel ID"` (merger matches BKF↔creative on these)
- Phase-4 output columns: `Creative File Name`, `Play Instructions`,
  `Start Date (MI)`, `End Date (MI)`, `Booking Start/End Date (BKF)`, `Comments`.

When the BKF template changes, update the `COL_*` / `BKF_*` / `MI_*` constants, not the
data.

## Conventions

- Canonical date format the pipeline aims for is `YYYY/MM/DD` (see `mi_pre_normaliser`).
- Canonical dimension is `WxH`, no spaces (e.g. `1080x1920`).
- Rotation/SOV values are plain numbers, never with a `%` sign; should sum to 100.
- Partial fill = a panel intentionally left with a blank `Creative File Name` because the
  MI only covered some dimensions. **Blank creative is valid and must still be uploaded** —
  it is a WARNING, not an error.
- Do-not-book = panels excluded from the final BKF (passed via `--do-not-book` or a flag
  column; see `DNB_FLAG_COLUMN`).

## Known issues / config TODOs (found during testing — fix carefully)

1. **Phase 3.5 crashes** when the MI has both `Creative Name` and `Creative File Name`
   (duplicate canonical columns → `_apply_normaliser` does `int(Series)`), and the
   `"dimension"` synonym mislabels `Pixel dimensions (Width)`/`(Height)`.
2. **Mixed date formats break Phase 5.** Phase 4.5 can emit both `YYYY/MM/DD` (MI) and
   `DD/MM/YYYY` (BKF) in the same column; Phase 5 parses with `dayfirst=True` over the
   whole column, silently coercing the minority to `NaT` → those rows lose their sort key
   and burst-week color. Normalise dates to one format before sorting.
3. **Partial-fill Play Instructions flagged as ERROR.** Phase 4.5 overwrites Play
   Instructions with the blank MI value; Phase 5 then errors on blank → blocks upload.
   Should be WARNING (match the Creative File Name handling).
4. `KNOWN_VALID_FORMATS` in Phase 3 holds none of the real Veridooh formats (e.g.
   `Portrait Retail`, `Rail Platform/Concourse`), so every panel flags "manual review".
5. Metabase is unconfigured: set `METABASE_URL`, `METABASE_TOKEN`, `METABASE_QUESTION_ID`
   and the `COL_*` names in [utils/metabase_client.py](utils/metabase_client.py);
   `prefetch()` is a per-row placeholder.
6. Phase 6 (Google Sheets + Monday.com) is not implemented.
