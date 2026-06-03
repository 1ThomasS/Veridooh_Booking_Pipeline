# Veridooh BFK Pipeline Orchestrator

An automated, event-driven pre-ingestion pipeline that chains Veridooh's
booking form preparation steps end-to-end.

---

## Architecture

```
Campaign inputs (IO + MI + Live Creative List)
        │
        ▼
Phase 1–2  io_to_bkf.py          ← existing (bkf-workflow skill)
        │
        ▼
Phase 3    format_resolver.py     ← NEW — Metabase lookup + WxH validation
        │
        ▼
Phase 4    mi-data-automation     ← existing (Claude skill, run separately)
        │
        ▼
Phase 4.5  phase45_merger.py      ← NEW — do-not-book filter + partial fill
        │
        ▼
Phase 5    phase5_postprocessor.py ← NEW — sort, burst highlight, compliance
        │
        ▼
Phase 6    upload_notify          ← FUTURE — Google Sheets + Monday.com API
        │
        ▼
DB · Monday.com auto-bot
```

---

## File Structure

```
veridooh_pipeline/
├── orchestrator.py             Main entry point — chains all phases
├── phase3_format_resolver.py   Format resolution via Metabase + WxH check
├── phase45_merger.py           BKF + creative merge, do-not-book, partial fill
├── phase5_postprocessor.py     Sort, burst highlight, compliance, Excel output
└── utils/
    ├── __init__.py
    └── metabase_client.py      Metabase query wrapper (configure before use)
```

---

## Setup

```bash
pip install pandas openpyxl requests
```

### Configure Metabase

Set environment variables before running:

```bash
export METABASE_URL="https://metabase.veridooh.com"
export METABASE_TOKEN="<your-session-token>"
export METABASE_QUESTION_ID="<question-id-from-url>"
```

Find the question ID by opening your panel lookup query in Metabase and
copying the number from the URL: `/question/<ID>`.

---

## Usage

### Full pipeline (most common)

```bash
# Step 1: Run Phases 1–2 + 3 (produces a phase3.csv, pauses for Phase 4)
python orchestrator.py \
    --campaign "JCD_Lifeblood" \
    --io "JCD_Lifeblood_IO.xlsx" \
    --mi "JCD_Lifeblood_MI.xlsx"

# Step 2: Run Phase 4 (mi-data-automation via Claude), save Output.csv

# Step 3: Resume with Phase 4.5 → 5
python orchestrator.py \
    --campaign "JCD_Lifeblood" \
    --bkf-draft "JCD_Lifeblood_phase3.csv" \
    --creative "Output.csv"
```

### Skip Phase 1+2 (BKF draft already exists)

```bash
python orchestrator.py \
    --campaign "CBA_oOh" \
    --bkf-draft "CBA_oOh_BKF.csv" \
    --creative "Output.csv" \
    --do-not-book P001 P045 P112
```

### Dry run (no Metabase queries)

```bash
python orchestrator.py \
    --campaign "Test" \
    --bkf-draft "draft.csv" \
    --creative "Output.csv" \
    --dry-run
```

### Run individual phases

```bash
# Phase 3 only
python phase3_format_resolver.py draft.csv --out draft_phase3.csv

# Phase 4.5 only
python phase45_merger.py draft_phase3.csv Output.csv --do-not-book P001 P045

# Phase 5 only
python phase5_postprocessor.py campaign_phase45.csv --out campaign_FINAL.xlsx
```

---

## Configuration TODOs

Before using in production, update these:

| File | Line | What to set |
|------|------|-------------|
| `utils/metabase_client.py` | `PANEL_LOOKUP_QUESTION_ID` | Metabase question ID |
| `utils/metabase_client.py` | `COL_PANEL_ID`, `COL_FORMAT`, etc. | Actual Metabase column names |
| `phase45_merger.py` | `DNB_FLAG_COLUMN`, `DNB_FLAG_VALUE` | How do-not-book panels are flagged in the IO |
| `phase5_postprocessor.py` | `KNOWN_VALID_FORMATS` (in phase3) | Expand as new formats are encountered |

---

## Adding Phase 6 (Google Sheets + Monday.com)

Replace `Pipeline.run_phase6()` in `orchestrator.py` with:

```python
import gspread                          # pip install gspread
from monday import MondayClient         # pip install monday

def run_phase6(self, final_xlsx_path: Path) -> None:
    # Upload to Google Sheets
    gc    = gspread.service_account(filename="service_account.json")
    sh    = gc.open(f"{self.name} BKF")
    ws    = sh.worksheet("VMO")
    # ... upload logic

    # Post link to Monday.com
    client = MondayClient(token=os.getenv("MONDAY_TOKEN"))
    # ... update item column
```

---

## What this project demonstrates (for your portfolio)

- **Data engineering**: end-to-end ETL pipeline with schema validation, format
  normalisation, and audit trails at each stage
- **System integration**: Python ↔ Metabase ↔ Excel ↔ Google Sheets ↔ Monday.com
- **Production quality**: modular design, configurable, CLI-accessible, dry-run
  mode, structured error reporting
- **Domain knowledge**: built on deep understanding of Veridooh's OOH verification
  workflow — not a generic template
