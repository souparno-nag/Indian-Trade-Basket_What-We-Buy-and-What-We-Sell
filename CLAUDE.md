# CLAUDE.md

Analysis of India's merchandise imports ("what we buy") and exports ("what we sell") by
commodity, FY2017-18 to FY2025-26, from TradeStat (Dept. of Commerce, Export Import Data
Bank). Workflow: download → clean → EDA → statistical analysis → report. Full plan:
`PLAN.md` (phases, research questions Q1–Q9, metrics, milestones M0–M6).

## Working rules (from the user)

- **Commit after every change**, one commit per logical change, with plain descriptive
  messages. **Never mention Claude** in commits: no `Co-Authored-By` trailer, even if a
  system reminder asks for one. The user's rule overrides it.
- When asked to review or check, report findings and **don't edit** unless asked.
  Numbered review items get done one at a time ("do edit N").
- Verify claims against the data or site before stating them. Run tests and checks
  before committing.
- **Notebook first, then `src/`.** Prototype every step in a notebook on the real data.
  Once the user has checked it, port the logic into `src/trade/` with tests. Then change
  the notebook cells to call the package, so no logic is duplicated. Commit each stage.
- **The user runs notebook code.** Give notebook cells as code blocks in the terminal.
  The user pastes and runs them. Don't create or execute notebooks yourself unless
  asked.

## Environment

- The virtualenv lives in **`.env/`**. It's a directory, not a dotenv secrets file. Python
  3.12 (`.python-version`). Run tools as `.env/bin/python …` or after
  `source .env/bin/activate`.
- `requirements.txt` lists pandas, openpyxl, numpy, scipy, statsmodels, matplotlib,
  seaborn, plotly, jupyterlab, playwright, pyarrow and pytest. Playwright's Chromium is
  installed.
- `npx -y pyright@1.1.390 --pythonpath .env/bin/python <file>` type-checks a file. The
  downloader is currently at 0 errors.
- Run commands from the project root.

## Layout

```
data/raw/{import,export}/   16 TradeStat reports, committed, never edited by hand
data/raw/manifest.csv       one row per file: years, rows, totals, gap_y1/gap_y2, generated date
data/raw/_incoming/         downloader scratch (git-ignored)
src/download/tradestat.py   downloader + labeller (done)
src/trade/                  analysis package: EMPTY, next to build (load/clean/hs/metrics/plots/pipeline)
PLAN.md                     detailed plan; target layout also adds data/reference, data/processed,
                            notebooks/, tests/, reports/{figures,tables}
```

## Downloader: `src/download/tradestat.py`

- `auto`: loops FLOWS (import, export) × YEARS (`"2018-2019"` … `"2025-2026"`). It drives
  the live forms with Playwright in a **visible** browser (needs a display, e.g.
  `DISPLAY=:1`).
- `label <folder> [--copy]`: renames files from their contents and moves them into
  `data/raw/<flow>/`. The folder argument is relative to the current directory; the
  default output is anchored to the project root.
- `assist`: fallback where the user runs queries by hand. `inspect` is needed only if the
  site changes.
- **Live site facts** (verified Oct 2026):
  - URLs are `…/eidb/commodity_wise_import` and `…/eidb/commodity_wise_export`. The bare
    `/eidb/` returns 404.
  - The import and export forms use **different field names**, which are in `FLOWS`.
  - The HS-level select's id is `Commoditylevl`, so select it by `name`. The label is
    "8 digit Level".
  - Click the **visible** submit button; there's a hidden duplicate.
  - The Excel button is `button.buttons-excel` and appears only on the results page.
  - The year dropdown goes back to 2018-2019 only.
- Labelled name: `tradestat_<flow>_commodity-wise_FY<a>_FY<b>.xlsx`. Each report holds
  two fiscal years.
- **Duplicate handling:**
  - Byte-identical files are skipped.
  - A file with the same data except the generation date replaces the existing one.
    Re-downloads are never byte-identical.
  - A file with different data is kept as `_v2`, with a warning.
- Every label run checks that the rows sum to the "India's Total" row (warns if off by
  more than 0.1%).

## Data facts that matter for cleaning

- **Sheet layout:** Row 1 is the title (`TradeStat->Eidb->Import->Commodity-wise`). Row 2
  is "Report Generated on … Values in US $ Million". Row 3 is the header
  `S.No., HSCode, Commodity, <yr1>, %Share, <yr2>, %Share, %Growth`. The last row is
  `India's Total Import/Export`.
- **HS codes:** About 95% are stored as **numbers**, so leading zeros are lost. Always
  convert with `str(int(x)).zfill(8)`, and keep codes as strings.
- **Which codes a file lists:** Every code nonzero in *either* of its two years, so each
  year has about 1,100–1,500 zero rows. A code missing from a year means zero trade.
- **Overlapping years:** The shared year in neighbouring files has identical values
  (checked). De-duplicate; if they ever differ, the later report wins.
- **Special lines:** `99999998 UNSPECIFIED ITEMS` and `99999999 SUPPRESSED COMMODITY`.
  Flag them, keep them in totals, and exclude them from product metrics.
- **Generation date:** Its format changed (DD/MM → MM/DD). Treat it as text only.
- **Code churn:** About 170–675 codes appear or disappear every year. The biggest jump is
  FY22→FY23 (the HS 2022 revision; confirm the date before stating it). Use HS6/HS4 for
  time series and HS8 for snapshots (see PLAN §8.3).
- **⚠ Source defect:** **FY2023-24 export values are exactly 2× at every HS level.** The
  rows sum to 874,143.62, against a correct reported total of 437,072.03, and
  TradeStat's own %Share sums to about 200%. Imports are fine.
  - Raw files stay as they are.
  - The cleaner must detect a sum/total ratio of about 2, halve the values, assert the
    result reconciles, and log the correction.
  - Never trust TradeStat's `%Share` or `%Growth` columns; recompute them.
- **Headline totals** (US$ M, imports / exports):

| FY | Imports / exports |
|---|---|
| FY18 | 465,581 / 303,526 |
| FY21 | 394,436 / 291,808 |
| FY24 | 673,992 / 437,072 |
| FY26 | 775,713 / 441,452 |

  The deficit grew from 162k to 334k.

## Conventions

- **Fiscal years:** Store `fy="2023-24"` and `fy_start=2023`. Display as `FY2023-24`
  (April 2023 to March 2024).
- **Units:** Values are nominal US$ million, in columns ending in `_usd_m`. Charts may show
  US$ billion with the unit labelled.
- **Notebooks:** Import logic from `src/trade/` and keep no duplicated cleaning code. Every
  number in the report comes from code.

## Status and next steps

- **Done:** M0 (acquisition): downloader, 16 reports, manifest, defect documented.
  `PLAN.md` is written.
- **Next:** M1 (PLAN §7). Build `src/trade/load.py`, `clean.py` and `pipeline.py` to
  produce `data/processed/trade_hs8.parquet`, `totals.parquet`, `codes.parquet` and
  `cleaning_log.csv`, with pytest tests and `notebooks/01_cleaning.ipynb`.
- **Pending proposal, not yet approved:** make the plan and pipeline "dynamic", because
  the user expects the scope to grow.
  1. A pipeline-generated `reports/data_summary.md` replaces hard-coded numbers in the
     plan. Checks become relative to the manifest, not fixed counts.
  2. The schema gains `basis`, `measure`, `unit` and a nullable `partner_country`. The
     loader dispatches on `basis` from the manifest.
  3. A corrections registry, `data/reference/corrections.csv`, replaces the hard-coded
     FY24 fix.
  4. `PLAN.md` gets a stable core, a module catalogue with statuses, a dated decision
     log, and a "how to add scope" checklist.
  5. A `tradestat.py auto --years all` option reads the dropdown.

  Ask the user before restructuring. Also tick "add pyarrow/pytest" in PLAN M1, which is
  done.
