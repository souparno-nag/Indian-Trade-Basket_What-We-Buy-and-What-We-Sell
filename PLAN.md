# Project Plan — Indian Trade Basket: What We Buy and What We Sell

A detailed, phase-by-phase plan for analysing India's merchandise imports and exports
by commodity, using official TradeStat (Export Import Data Bank, Department of Commerce)
data and the Python data stack.

Status markers used below: `[x]` done, `[ ]` to do, `[~]` optional / stretch.

---

## Table of contents

1. [Goals and research questions](#1-goals-and-research-questions)
2. [Where the project stands](#2-where-the-project-stands)
3. [The data: inventory, facts and quirks](#3-the-data-inventory-facts-and-quirks)
4. [Repository layout](#4-repository-layout)
5. [Environment and tooling](#5-environment-and-tooling)
6. [Phase 0 — Data acquisition](#6-phase-0--data-acquisition)
7. [Phase 1 — Loading and cleaning](#7-phase-1--loading-and-cleaning)
8. [Phase 2 — Enrichment and reference data](#8-phase-2--enrichment-and-reference-data)
9. [Phase 3 — Exploratory data analysis](#9-phase-3--exploratory-data-analysis)
10. [Phase 4 — Statistical analysis](#10-phase-4--statistical-analysis)
11. [Phase 5 — Joint import–export analysis](#11-phase-5--joint-importexport-analysis)
12. [Phase 6 — Reporting](#12-phase-6--reporting)
13. [Testing and validation strategy](#13-testing-and-validation-strategy)
14. [Conventions](#14-conventions)
15. [Risks, limitations and how we handle them](#15-risks-limitations-and-how-we-handle-them)
16. [Extensions](#16-extensions)
17. [Open decisions](#17-open-decisions)
18. [Milestone checklist](#18-milestone-checklist)

---

## 1. Goals and research questions

### 1.1 Overall goal

Describe, measure and explain how India's **import basket** (what we buy) and **export
basket** (what we sell) have changed between **FY2017-18 and FY2025-26**, at commodity
level (8-digit HS), with reproducible Python code: raw files → clean data → EDA →
statistical analysis → written findings.

### 1.2 Research questions

Each question names the phase that answers it.

**Composition: what do we trade?**
- Q1. What are the largest import and export commodities and chapters each year, and how
  have their ranks and shares changed? (Phase 3)
- Q2. How concentrated is each basket? Do a handful of products dominate, and is that
  concentration rising or falling? (Phases 3, 4)
- Q3. Which sectors (energy, gold and gems, electronics, machinery, chemicals, pharma,
  agriculture, textiles, vehicles, …) make up each basket, and how have their shares
  moved? (Phases 2, 3)

**Change: how is the basket shifting?**
- Q4. Which commodities grew or shrank fastest, in absolute dollars and in percentage
  terms, over the whole period (CAGR) and year on year? (Phase 4)
- Q5. How much structural change happened from one year to the next, and were some years
  (e.g. FY2020-21, FY2022-23) unusually disruptive? (Phase 4)
- Q6. Is the export basket diversifying (more products, less concentration), and is
  growth coming from new products (extensive margin) or from bigger sales of existing
  ones (intensive margin)? (Phase 4)

**Balance: how do buying and selling relate?**
- Q7. Where is India a net importer and where a net exporter, by chapter and by product?
  Which deficits are widening? (Phase 5)
- Q8. How much trade is intra-industry, meaning India both imports and exports the same
  product, for example rough versus cut diamonds, or crude oil versus refined fuels?
  (Phase 5)
- Q9. Do imports and exports move together within a sector, which would suggest
  import-dependent export production? (Phase 5)

**Episodes worth testing** (hypotheses to check against the data, not assumed facts):
- H1. FY2020-21 shows a broad COVID-19 contraction followed by a sharp FY2021-22 rebound.
- H2. Electronics exports, especially smartphones (HS 85171300), grew structurally after
  FY2020-21.
- H3. Crude oil and gold dominate import value, so import totals swing with their prices.
- H4. Petroleum products and cut diamonds show up on both sides of the ledger
  (import-to-process-to-export chains).

---

## 2. Where the project stands

As of 2026-10-26:

- [x] Git repository set up, with one commit per change.
- [x] Python 3.12 virtual environment at `.env/`, pinned in `.python-version`.
- [x] `requirements.txt` lists pandas, openpyxl, numpy, scipy, statsmodels, matplotlib,
      seaborn, plotly, jupyterlab and playwright.
- [x] Downloader `src/download/tradestat.py` handles automated download, labelling,
      manifest, total checks and duplicate handling.
- [x] Raw data downloaded and committed: 16 reports (8 import, 8 export), FY2017-18 to
      FY2025-26.
- [x] One source-data defect found and documented (FY2023-24 exports doubled; see §3.4).
- [ ] No loader, cleaner, notebooks, tests, README or analysis yet. `src/trade/` exists
      but is empty.

---

## 3. The data: inventory, facts and quirks

### 3.1 Source

- **Portal:** TradeStat, Export Import Data Bank (EIDB), Department of Commerce, Government
  of India.
  - Imports: <https://tradestat.commerce.gov.in/eidb/commodity_wise_import>
  - Exports: <https://tradestat.commerce.gov.in/eidb/commodity_wise_export>
- **Report:** "Commodity-wise", all commodities, **8-digit HS level**, **US $ Million**.
- **Underlying compiler:** DGCI&S (Directorate General of Commercial Intelligence and
  Statistics), Kolkata.
- **Site notes** (taken from the portal page):
  - Imports and exports include re-imports and re-exports.
  - US$ figures from 2006-07 onwards are sums of monthly US$ figures.
  - A year label like "2012-2013" means April 2012 to March 2013.

### 3.2 Files

Location: `data/raw/{import,export}/tradestat_<flow>_commodity-wise_FY<a>_FY<b>.xlsx`.

- **Two years per file:** Each report covers two consecutive fiscal years, the selected
  year and the one before it. Eight reports per flow therefore cover nine fiscal years,
  FY2017-18 to FY2025-26.
- **Overlap checked:** The overlapping year in neighbouring files was checked on
  2026-10-26 and the values are **identical**. No revisions differ between files from the
  same download date.
- **`data/raw/manifest.csv`:** One row per file, with flow, years, row count, reported
  totals, `gap_y1`/`gap_y2` (reported total minus the sum of the rows) and the generation
  date.

### 3.3 Layout of each sheet

| Row | Content |
|---|---|
| 1 | Title: `TradeStat->Eidb->Import->Commodity-wise` (or `Export`) |
| 2 | `Report Generated on: <date> - Values in US $ Million` |
| 3 | Header: `S.No., HSCode, Commodity, <year1>, %Share, <year2>, %Share, %Growth` |
| 4 … n-1 | One row per 8-digit HS code |
| n | `India's Total Import` / `India's Total Export`, with totals in the year columns |

**Facts measured on the downloaded data:**
- **Codes per year:** About 10,260–10,850 import codes and 10,940–11,560 export codes.
- **Codes stored as numbers:** About 95% of HS codes are numeric Excel cells, not text, so
  **leading zeros are lost** (e.g. `1012100` instead of `01012100`). Only some codes are
  stored as text with the zero kept. Every code must be left-padded to 8 characters.
- **Which codes a file lists:** A file includes every code that is nonzero in *either* of
  its two years. That gives about 1,060–1,530 zero-valued rows per year, and codes
  missing from a year mean "no trade recorded" (zero).
- **No blanks or negatives:** No negative values and no blank value cells were found in
  the value columns.
- **Blank growth:** `%Growth` is blank where the earlier year is zero, because the growth
  rate is undefined there.
- **Special codes:** `99999998 UNSPECIFIED ITEMS` and `99999999 SUPPRESSED COMMODITY`.
  These are not real products and need flagging, not mixing with real trade lines.
- **Date format change:** The generation date switched from DD/MM to MM/DD between
  downloads. Treat it as text only.

**Headline totals** (US$ million, taken from the "India's Total" rows):

| FY | Imports | Exports | Balance | Exports / imports |
|---|---:|---:|---:|---:|
| 2017-18 | 465,581 | 303,526 | −162,055 | 65.2% |
| 2018-19 | 514,078 | 330,078 | −184,000 | 64.2% |
| 2019-20 | 474,709 | 313,361 | −161,348 | 66.0% |
| 2020-21 | 394,436 | 291,808 | −102,627 | 74.0% |
| 2021-22 | 613,052 | 422,004 | −191,048 | 68.8% |
| 2022-23 | 715,969 | 451,070 | −264,899 | 63.0% |
| 2023-24 | 673,992 | 437,072 | −236,920 | 64.8% |
| 2024-25 | 721,200 | 437,705 | −283,496 | 60.7% |
| 2025-26 | 775,713 | 441,452 | −334,261 | 56.9% |

**Largest chapters in FY2025-26:**
- **Imports:**
  - Ch 27 mineral fuels: 26.2%
  - Ch 71 gold, gems and jewellery: 14.1%
  - Ch 85 electrical and electronics: 13.5%
  - Ch 84 machinery: 9.5%
  - The top 20 codes make up 47.3% of all imports.
- **Exports:**
  - Ch 27: 12.7%
  - Ch 85: 12.2%
  - Ch 84: 8.3%
  - Ch 71: 6.4%
  - Ch 87 vehicles: 5.9%
  - Ch 30 pharma: 5.7%
  - The top 20 codes make up 32.6% of all exports.

### 3.4 Known source defect: FY2023-24 exports are doubled

- **What:** Every FY2023-24 **export** commodity value is exactly 2× what it should be,
  at every HS level (seen at 8-digit and 2-digit).
  - The rows sum to 874,143.62, but the reported total is 437,072.03.
  - The site's own `%Share` column sums to about 200%.
  - Its `%Growth` column was computed from the doubled values.
- **Correct figure:** The **reported total** is right; it matches India's official FY24
  merchandise exports of about $437B.
- **Scope:** FY2023-24 **imports** are fine.
- **Affected files:** `tradestat_export_commodity-wise_FY2022-23_FY2023-24.xlsx` (year-2
  column) and `tradestat_export_commodity-wise_FY2023-24_FY2024-25.xlsx` (year-1 column).
- **Handling:**
  - Raw files stay as downloaded.
  - The cleaner divides FY2023-24 export values by 2, then **asserts** that the halved sum
    matches the reported total to within rounding (halved sum 437,071.81 vs 437,072.03).
  - The cleaner never trusts TradeStat's `%Share` or `%Growth` columns, and recomputes
    them.
  - After any re-download, check again whether the defect is still there. TradeStat may
    fix it, and then the correction must turn itself off (see §7.3, step 9).

### 3.5 HS code churn

8-digit codes appear and disappear **every year**, not just when the HS nomenclature is
revised:

| Transition | Imports dropped / new | Exports dropped / new |
|---|---|---|
| FY19 → FY20 | 203 / 318 | 202 / 212 |
| FY20 → FY21 | 264 / 440 | 214 / 460 |
| FY21 → FY22 | 461 / 250 | 319 / 248 |
| FY22 → FY23 | 279 / **675** | 169 / **580** |
| FY23 → FY24 | 512 / 541 | 475 / 472 |
| FY24 → FY25 | 396 / 308 | 313 / 212 |
| FY25 → FY26 | 250 / 419 | 194 / 322 |

- **Two causes:** Some churn is real (products start or stop trading, so they enter or
  leave the file). Some is reclassification.
- **HS 2022:** The spike in new codes in FY2022-23 fits India adopting the HS 2022
  nomenclature in its ITC-HS tariff schedule. Confirm the exact effective date from an
  official source before stating it in the report.
- **National lines:** Digits 7–8 are national tariff lines, which India can amend in any
  budget. So 8-digit series are not stable over time, and time-series work needs care
  (§8.3).

---

## 4. Repository layout

Target structure. Folders that don't exist yet are created when their phase begins.

```
Indian-Trade-Basket_What-We-Buy-and-What-We-Sell/
├── PLAN.md                    # this file
├── README.md                  # [ ] project summary, setup, how to reproduce
├── requirements.txt
├── .python-version            # 3.12
├── .gitignore
├── .env/                      # virtualenv (git-ignored)
├── data/
│   ├── raw/                   # untouched downloads; never edited by hand
│   │   ├── import/            # 8 reports
│   │   ├── export/            # 8 reports
│   │   ├── manifest.csv
│   │   └── _incoming/         # downloader scratch (git-ignored)
│   ├── reference/             # [ ] hand-made lookup tables (HS sections, sector groups)
│   ├── interim/               # [ ] intermediate outputs, if any (git-ignored)
│   └── processed/             # [ ] tidy, analysis-ready tables
├── src/
│   ├── download/
│   │   └── tradestat.py       # downloader / labeller
│   └── trade/                 # [ ] reusable analysis package
│       ├── __init__.py
│       ├── load.py            # read raw xlsx → DataFrame
│       ├── clean.py           # cleaning, corrections, validation
│       ├── hs.py              # HS hierarchy, sections, sector groups, concordance
│       ├── metrics.py         # shares, growth, CAGR, HHI, Gini, GL index, …
│       ├── plots.py           # shared plotting helpers and style
│       └── pipeline.py        # one command: raw → processed
├── notebooks/
│   ├── 01_cleaning.ipynb      # walk-through and checks of the pipeline output
│   ├── 02_eda_imports.ipynb
│   ├── 03_eda_exports.ipynb
│   ├── 04_statistics.ipynb
│   ├── 05_balance_and_iit.ipynb
│   └── 99_scratch.ipynb       # experiments (not relied upon)
├── tests/                     # [ ] pytest
│   ├── test_load.py
│   ├── test_clean.py
│   └── test_metrics.py
└── reports/
    ├── figures/               # exported charts (PNG/SVG/HTML)
    ├── tables/                # exported summary tables (CSV)
    └── report.md              # final write-up
```

**Rules:**
- **Notebooks import from `src/trade/`.** Logic used in more than one place moves into
  the package. Notebooks narrate and plot; they hold no duplicated cleaning code.
- **One way in for processed data:** `data/processed/` is produced only by
  `python -m trade.pipeline` (or a matching script). Never edit it by hand.
- **Git:** Processed outputs can be regenerated, so ignore them in git if they grow
  large. They're small for this project (about 200k rows), so committing a parquet
  snapshot is also acceptable (see §17).

---

## 5. Environment and tooling

- **Python 3.12** in `.env/`. Activate it with `source .env/bin/activate`.
- **Already in `requirements.txt`:** pandas, openpyxl, numpy, scipy, statsmodels,
  matplotlib, seaborn, plotly, jupyterlab, playwright.
- **To add when the phase needs them:**
  - `pyarrow`: parquet I/O for processed data (Phase 1).
  - `pytest`: tests (Phase 1).
  - `scikit-learn`: clustering and scaling (Phase 4).
  - `[~] pymannkendall`: Mann–Kendall trend test (Phase 4). Hand-writing it is the
    alternative.
  - `[~] powerlaw`: heavy-tail fitting (Phase 4).
  - `[~] squarify`: matplotlib treemaps, unless plotly treemaps are enough.
- **Pin versions:** Once the environment is stable, run
  `pip freeze > requirements.lock` so results can be reproduced exactly.
- **Notebook hygiene:** Clear outputs before committing, or use `nbstripout`, to keep
  diffs readable.

---

## 6. Phase 0 — Data acquisition

**Status: done.** This section records how to repeat or extend it.

- [x] `python src/download/tradestat.py auto` downloads FLOWS × YEARS (2 × 8) into
      `data/raw/{flow}/`.
- [x] Each file is labelled from its contents (flow and years read from inside the
      sheet), logged in `manifest.csv`, and checked against its reported total. Warnings
      start with `!`.
- [x] Re-downloads with identical data replace files in place. Revised figures are kept
      as `_v2`.
- [ ] **Refresh routine** (when TradeStat publishes new or revised data):
  1. Run `auto`.
  2. Read every `!` line in the output.
  3. Inspect any `_v2` file.
  4. Check whether the FY2023-24 export defect still exists (`gap_y*` in the manifest).
  5. Commit.
- [~] Extensions: quantity reports, ₹ crore reports, country-wise reports (§16).

---

## 7. Phase 1 — Loading and cleaning

**Goal:** One tidy, validated table holding every flow × fiscal year × HS8 value, with
the known defects corrected, plus a log of every correction.

### 7.1 Target schema: `data/processed/trade_hs8.parquet`

Long format, one row per (flow, fiscal year, HS8 code):

| Column | Type | Description |
|---|---|---|
| `flow` | category | `import` / `export` |
| `fy` | str | `"2017-18"` … `"2025-26"` |
| `fy_start` | int16 | 2017 … 2025 (for sorting and regression) |
| `hs8` | str(8) | zero-padded code, e.g. `"01012100"` |
| `hs6`, `hs4`, `hs2` | str | prefixes of `hs8` |
| `description` | str | cleaned description (§7.3, step 7) |
| `value_usd_m` | float64 | value in US$ million, corrected |
| `value_raw_usd_m` | float64 | value as published (before corrections) |
| `is_special` | bool | `99999998` / `99999999` and any other non-product lines |
| `correction` | str / null | e.g. `"fy2023-24_export_halved"` |
| `source_file` | str | raw file the value came from |

**Companion tables:**
- **`data/processed/totals.parquet`:** Reported "India's Total" per flow and year, with
  the sum of rows before and after correction.
- **`data/processed/codes.parquet`:** One row per HS8 code, with its latest description,
  first and last year seen, the years present, and whether it is in the stable panel
  (§8.3).
- **`data/processed/cleaning_log.csv`:** Every correction or exclusion applied, with
  counts.

### 7.2 `src/trade/load.py`

1. `read_report(path) -> (meta, DataFrame)`:
   - Read with `pandas.read_excel(path, header=2, dtype={"HSCode": str})`. Openpyxl
     still returns numbers for numeric cells, so convert each value with `str(int(x))`
     where it's numeric, then `.zfill(8)`.
   - Rename the columns to `sno, hs8, description, y1, share1, y2, share2, growth`.
   - Take the two fiscal-year labels from the header row (`"2023 - 2024"` →
     `"2023-24"`).
   - Separate the `India's Total …` row into `meta` (and keep it for validation).
   - Take the flow from the title cell (row 1), not the file name.
2. `read_all(raw_dir) -> DataFrame`: read every report, melt `y1`/`y2` into long rows,
   and attach `source_file`.
3. **De-duplicate the overlapping years.** Each fiscal year except the first and last
   appears in two files.
   - Assert that the two copies agree (currently they match exactly). If they don't, keep
     the value from the **later** report, which reflects newer data, and log every
     difference.
   - After de-duplication, a code is absent for a year if neither file lists it, which
     means zero trade.

### 7.3 `src/trade/clean.py`: steps, in order

1. **Code hygiene:**
   - Every `hs8` must match `^\d{8}$`. Rows that don't are logged and dropped (none
     expected).
   - Assert there are no duplicate `(flow, fy, hs8)` rows after de-duplication.
2. **Special lines:** Set `is_special=True` for `99999998` and `99999999`, plus any other
   code starting with `9999` or whose description contains `UNSPECIFIED` or
   `SUPPRESSED`. Keep these rows so totals still reconcile, but leave them out of
   product-level rankings and metrics by default.
3. **Value types:** Coerce to float, and treat blanks as 0.0. Assert no negatives.
4. **Recompute derived columns.** Drop TradeStat's `%Share` and `%Growth`, because they're
   wrong for FY2023-24 exports and redundant otherwise. Recompute both in `metrics.py`.
5. **Zero rows:**
   - Keep explicit zeros as listed.
   - For a balanced panel, do **not** fill in every code for every year at this stage.
     `metrics.py` builds the code × year panel on demand, filling missing values with 0.
6. **Units:** All values are US$ million. Record this in column names (`_usd_m`) and in
   the README.
7. **Descriptions:**
   - Normalise whitespace (collapse repeated spaces, which are common, e.g.
     `"OTHER  MEDCNE"`) and strip.
   - Keep the original case, since descriptions mix upper case and "Smartphones".
   - Use the description from the **latest** year a code appears in.
   - Keep a `description_variants` count so codes whose meaning may have changed can be
     spotted.
8. **Reconcile with totals.** For every (flow, fy), compare the sum of rows (including
   special lines) to the reported total.
   - Expect the gap to be ≤ 0.1% (observed: ≤ about 1.5 US$ million, from rounding).
   - Fail loudly otherwise, unless step 9 explains the gap.
9. **FY2023-24 export doubling correction:**
   - Detect rather than hard-code. If `sum(rows) / reported_total` is between 1.99 and
     2.01 for a (flow, fy), halve that year's values for that flow, set
     `correction="<fy>_<flow>_halved"`, keep `value_raw_usd_m`, and re-run step 8.
   - Today this triggers only for exports, FY2023-24. If TradeStat fixes the data, the
     ratio becomes 1.0 and the correction switches itself off.
   - The detection rule is a safeguard. Also emit a warning listing every (flow, fy) it
     changed.
10. **Write outputs** (§7.1) and the cleaning log. Print a short summary: rows per flow ×
    year, corrections, specials excluded, and gaps.

### 7.4 `src/trade/pipeline.py`

- `python -m trade.pipeline` runs load → clean → enrich (Phase 2) → write, with no
  hidden notebook state.
- It must finish quickly. Reading 16 xlsx files of about 11k rows each is the slow part,
  so cache an interim parquet of the raw long table if needed.

### 7.5 Phase 1 exit criteria

- [ ] `trade_hs8.parquet`, `totals.parquet`, `codes.parquet` and `cleaning_log.csv`
      exist.
- [ ] For every (flow, fy), the corrected sum of rows is within 0.1% of the reported
      total.
- [ ] FY2023-24 export total after correction ≈ 437,072, and the correction is logged.
- [ ] Tests in `tests/test_load.py` and `tests/test_clean.py` pass.
- [ ] `notebooks/01_cleaning.ipynb` shows the checks with their outputs.

---

## 8. Phase 2 — Enrichment and reference data

### 8.1 HS hierarchy

The HS2, HS4 and HS6 prefixes come straight from `hs8`. Chapter (HS2) names come from a
reference table. Write `data/reference/hs_chapters.csv` (98 chapters, plus chapter 77,
which is reserved) by hand or from an official HS list, with columns:
`hs2, chapter_name, section`.

### 8.2 HS sections and analytical sector groups

- **`data/reference/hs_sections.csv`:** The 21 official HS sections (I–XXI), each with a
  chapter range (e.g. Section V *Mineral products*: chapters 25–27).
- **`data/reference/sector_groups.csv`:** A simpler grouping for storytelling, defined in
  this project and documented as such:

| Group | HS chapters (initial proposal) |
|---|---|
| Energy (crude, petroleum products, coal, gas) | 27 |
| Gold, silver, gems and jewellery | 71 |
| Electronics and electrical | 85 |
| Machinery and mechanical appliances | 84 |
| Chemicals (organic/inorganic) | 28, 29, 38 |
| Pharmaceuticals | 30 |
| Plastics and rubber | 39, 40 |
| Metals (iron, steel, base metals) | 72–83 |
| Vehicles and transport equipment | 86–89 |
| Textiles and apparel | 50–63 |
| Agriculture and food | 01–24 |
| Minerals and ores (excluding fuels) | 25, 26 |
| Instruments (optical, medical, measuring) | 90–92 |
| Other manufactures | everything else |

- **Finer splits:** The groups are chosen at chapter level first. Where the story needs
  it, split at HS4. For example, split 2709 (crude) from 2710 (refined products); within
  71, split 7102 (diamonds), 7108 (gold) and 7113 (jewellery); and pull out 851713
  (smartphones).
- **Ownership:** The groupings are judgement calls. Keep them in CSV, not in code, so
  they can be reviewed and changed.

### 8.3 Handling code churn (concordance strategy)

8-digit codes change every year (§3.5), so choose the level of analysis by the question:

- **Single-year snapshots** (rankings, concentration within a year): use HS8 directly.
- **Time series of individual products:**
  - Default to **HS6**. It is internationally harmonised and changes only with HS
    revisions (2017 → 2022).
  - Use **HS4** for long series that cross the FY2022-23 revision.
- **Stable panel:**
  - Define it as the HS8 codes with nonzero trade in all nine years, separately per
    flow.
  - Report the share of total value it covers each year. If it covers most of the value,
    analysis restricted to it is reasonably representative.
  - Use it for HS8-level growth comparisons.
- **[~] HS 2017 → HS 2022 concordance:**
  - The WCO and UN Statistics Division publish correlation tables between HS editions at
    6 digits.
  - Download the table into `data/reference/` and map pre-revision HS6 codes to HS 2022.
  - Split one-to-many mappings by value share (or flag them, not split).
  - Only worth it if HS6 series around FY2022-23 look broken.
- **Churn analysis:** Treat churn as an analysis topic too (extensive margin, §10.5).

### 8.4 Phase 2 exit criteria

- [ ] The three reference CSVs exist and are documented.
- [ ] Every `hs2` in the data maps to a chapter, a section and a sector group. Assert that
      none are unmapped.
- [ ] `codes.parquet` has `in_stable_panel_import` and `in_stable_panel_export` flags, and
      the coverage table exists.

---

## 9. Phase 3 — Exploratory data analysis

Run separately for imports (`02_eda_imports.ipynb`) and exports (`03_eda_exports.ipynb`),
using the same functions. Special lines are excluded from product-level views and
included in totals.

### 9.1 Totals and trends

- [ ] Annual totals for both flows, FY2017-18 to FY2025-26, with YoY growth. Mark
      FY2020-21.
- [ ] Trade balance and export/import coverage ratio over time (the table in §3.3 as a
      chart).

### 9.2 Composition

- [ ] Top 10 and top 20 HS8 products per year: value, share and rank.
- [ ] Top chapters and sector groups per year. Show them as a stacked area or 100%
      stacked bar chart of shares over time, and as a treemap for FY2025-26
      (section → chapter → HS4).
- [ ] **Rank dynamics:** a bump chart of the top 15 HS4 groups across years, and a
      "who entered or left the top 20" table.
- [ ] **Two-sided view:** for each sector group, import share vs export share in the
      first and last years.

### 9.3 Distribution of values

- [ ] Histogram or KDE of `log10(value_usd_m)` for nonzero lines per year. The values are
      extremely skewed (crude oil alone is about 17% of imports).
- [ ] A rank-size (Zipf) plot on log–log axes, per flow and year.
- [ ] Count of lines per value bucket (< 0.1, 0.1–1, 1–10, 10–100, 100–1,000,
      > 1,000 US$ million), and how the value total splits across those buckets.

### 9.4 Concentration (descriptive)

- [ ] **Pareto curve:** cumulative share of value against cumulative share of product
      lines.
- [ ] Top-N shares (N = 1, 5, 10, 20, 50, 100) over time.
- [ ] A Lorenz curve per year with the Gini coefficient (formal treatment in §10.3).

### 9.5 Movers

- [ ] Biggest absolute gainers and losers, FY2017-18 → FY2025-26 and year on year
      (US$ million).
- [ ] Biggest percentage gainers and losers among lines above a size threshold (e.g.
      ≥ 50 US$ million in the base year), so tiny bases don't dominate.
- [ ] New and vanished lines per year, with their combined value.

### 9.6 Data-quality EDA (sanity checks that feed back into Phase 1)

- [ ] Codes whose value jumps more than 10× or falls below a tenth year on year. Check
      whether these are reclassification artefacts, real shocks, or data errors.
- [ ] Codes whose description changes between years.
- [ ] FY2023-24 exports after correction: check its share pattern against FY2022-23 and
      FY2024-25 for continuity.

### 9.7 Visual conventions

- **Style:** One shared style in `src/trade/plots.py`, with consistent colours per sector
  group across all charts. One colour means one sector everywhere.
- **Labels:** Always label units (US$ billion on charts, converted from million, for
  readability) and fiscal years in `FY2023-24` form.
- **Export:**
  - Save static PNG/SVG to `reports/figures/` with descriptive file names, e.g.
    `imports_sector_shares_fy18_fy26.png`.
  - Interactive plotly versions are optional, saved as HTML.
- **Log scales:** Use them for distributions, and state when a scale is logarithmic.

---

## 10. Phase 4 — Statistical analysis

All metrics live in `src/trade/metrics.py` with unit tests on small hand-checked
examples. Each is computed per flow, per year, and at a stated HS level.

### 10.1 Growth

- [ ] **YoY growth:** `g_t = v_t / v_{t-1} − 1`. Also log growth `Δln v`, which
      aggregates better.
- [ ] **CAGR:** `(v_end / v_start)^(1/n) − 1`, with n = 8 for FY18 → FY26. Undefined
      when `v_start = 0`; report those cases as "new" lines instead.
- [ ] **Robust summaries:** median and interquartile range of growth across lines, and
      value-weighted mean growth.
- [ ] **Contribution to growth:** a line's contribution to total growth is
      `(v_t − v_{t-1}) / Total_{t-1}`. Rank by contribution, not by growth rate.
- [ ] **Shift-share decomposition** of total change into sector-group contributions.

### 10.2 Volatility and stability

- [ ] Standard deviation of log growth per line or sector across years (lines with ≥ 6
      nonzero years).
- [ ] **Rank stability:** Spearman rank correlation of product values between
      consecutive years and between FY18 and FY26.
- [ ] **Persistence:** the share of the top-50 products in year t still in the top 50 in
      t+1.

### 10.3 Concentration and diversification

For each flow and year, with shares `s_i` that sum to 1 across non-special lines:

- [ ] **HHI** = Σ s_i², at HS8, HS6, HS4 and HS2. Also report the "equivalent number of
      products", 1/HHI.
- [ ] **Gini coefficient** of line values.
- [ ] **Theil index**, split into between-sector and within-sector components. This
      answers whether concentration comes from the sector mix or from the products within
      sectors.
- [ ] **Number of active lines** (value > 0 and value > 1 US$ million) as a count-based
      diversification measure.
- [ ] **Trend tests:** a Mann–Kendall test and Sen's slope for each concentration series.
      State clearly that with n = 9 annual points the power is low, so treat these as
      descriptive evidence, not proof.

### 10.4 Structural change

- [ ] **Year-to-year dissimilarity of share vectors:**
  - Lawrence index, ½ Σ |s_i,t − s_i,t−1|
  - Cosine similarity
  - Jensen–Shannon divergence
- [ ] A **heat map of pairwise dissimilarity** across all nine years, which shows whether
      the basket drifts steadily or jumps.
- [ ] Highlight the transitions with the largest changes and break them down by sector,
      to see which sector drove each jump.
- [ ] Repeat at HS4 to separate real structural change from HS8 code churn.

### 10.5 Extensive vs intensive margin (exports first, then imports)

- [ ] Split export growth between consecutive years (and FY18 → FY26) into three parts:
  - **Intensive:** growth of lines traded in both years.
  - **Extensive, entry:** value of new lines.
  - **Extensive, exit:** value lost from lines that dropped out.
- [ ] Do this at HS6 as well as HS8, so pure reclassification inflates entries and exits
      less.

### 10.6 Distribution fitting

- [ ] Fit log-normal and power-law (Pareto tail) models to the value distribution of
      nonzero lines. Compare them with likelihood-ratio tests (Clauset–Shalizi–Newman
      method, via `powerlaw` or hand-written code), and check with QQ plots.
- [ ] Rank-size regression: `ln(rank) = a − b·ln(value)` for the upper tail. Track the
      slope `b` over time as a check on concentration.

### 10.7 Outliers and anomalies

- [ ] **Robust z-score** on log growth: `(x − median) / (1.4826·MAD)`. Flag |z| > 3.5
      for review.
- [ ] Cross-check flagged lines against the Phase 3 data-quality list. Separate real
      shocks from artefacts.

### 10.8 Clustering of trajectories

- [ ] Take HS4 groups above a value threshold. Build share trajectories across nine
      years, normalised to their own mean (shape) or z-scored.
- [ ] Run k-means (k chosen by silhouette score and inspection) and hierarchical
      clustering (Ward). Name the clusters by eye: "steady risers", "COVID-dip and
      rebound", "declining", "volatile".
- [ ] Show cluster-centre trajectories with member counts and value.

### 10.9 Inference: uncertainty and honesty

- [ ] **Bootstrap** (resampling product lines) confidence intervals for median growth,
      HHI changes and similar statistics, where line-level resampling makes sense.
- [ ] Distribution comparisons (e.g. growth before vs after FY2020-21) with Mann–Whitney
      / Wilcoxon tests, stating their assumptions.
- [ ] **No causal claims** from these data alone. Policy or price episodes (H1–H4) are
      described as consistent or inconsistent with the patterns, not proven.
- [ ] Values are **nominal US$**. A rising import bill can mean higher prices, a weaker
      rupee, or more volume. Don't interpret value changes as volume without quantity data
      (§16).

### 10.10 Phase 4 exit criteria

- [ ] `metrics.py` covers everything in §10.1–10.5, with tests.
- [ ] `04_statistics.ipynb` produces each table and figure. Key tables are exported to
      `reports/tables/`.
- [ ] A short written interpretation for each metric family, including caveats.

---

## 11. Phase 5 — Joint import–export analysis

### 11.1 Trade balance by product and sector

- [ ] Net balance (exports − imports) per HS2, HS4 and sector group per year. Rank the
      biggest deficits and surpluses, and the fastest-widening ones.
- [ ] A **waterfall chart** of the FY18 → FY26 change in the total deficit, by sector.
- [ ] **Coverage ratio** (exports ÷ imports) by sector over time.

### 11.2 Intra-industry trade (Grubel–Lloyd)

- [ ] For each product i: `GL_i = 1 − |X_i − M_i| / (X_i + M_i)`.
  - 0 means one-way trade; 1 means perfectly balanced two-way trade.
  - Compute at HS6 and HS4. HS8 is too fine, because national lines differ between the
    import and export schedules.
- [ ] **Aggregate GL** (value-weighted) per year, and per sector group.
- [ ] Products with large flows in both directions: check whether H4 (crude → refined
      products, rough → cut diamonds, gold → jewellery) shows up as high-GL or as linked
      HS4 pairs. Note: those chains cross HS4 codes (e.g. 2709 imported → 2710 exported),
      so also compare **linked pairs**, not just identical codes.

### 11.3 Co-movement

- [ ] Correlation of import growth and export growth within each sector group across
      years (n = 8 growth observations each, so treat as descriptive).
- [ ] Lagged relationships (do imports at t relate to exports at t+1?), as exploratory
      only, given the very short series.

### 11.4 Import dependence (exploratory)

- [ ] Sectors where imports exceed exports and imports are growing faster, as candidates
      for "dependence".
- [ ] Compare electronics (85) imports with electronics exports: is export growth matched
      by component imports such as ICs (8542)?

### 11.5 Phase 5 exit criteria

- [ ] `05_balance_and_iit.ipynb` with balance, GL and co-movement results.
- [ ] Linked-pair table (HS4 import → HS4 export) defined in `data/reference/` and used.

---

## 12. Phase 6 — Reporting

- [ ] `reports/report.md`, the final write-up, structured around Q1–Q9:
  1. Executive summary: 5–8 headline findings with numbers.
  2. Data and method: source, coverage, cleaning, the FY2023-24 correction, limitations.
  3. What we buy: composition, concentration, change.
  4. What we sell: composition, diversification, margins.
  5. The balance: deficits, intra-industry trade, linked chains.
  6. Statistical results with uncertainty.
  7. Limitations and further work.
  8. Appendix: metric definitions, sector-group mapping, data dictionary.
- [ ] Every number in the report is produced by code (exported tables), not typed by
      hand.
- [ ] `README.md`: purpose, setup (`python3.12 -m venv .env`,
      `pip install -r requirements.txt`, `playwright install chromium`), how to reproduce
      (download → pipeline → notebooks), repository map, data source and licence or
      attribution.
- [~] An HTML version of the report, or a small dashboard (plotly / a static site), if it
      will be shared.

---

## 13. Testing and validation strategy

**Unit tests (`tests/`, pytest):**
- `test_load.py`:
  - Parse a small fixture xlsx built in the test, with numeric and text codes, a total
    row and a blank growth cell.
  - Check that codes are zero-padded, years are parsed, the total is separated, and the
    melt is correct.
- `test_clean.py`:
  - Detect and halve a synthetic doubled year.
  - Do *not* halve a normal year.
  - Fail on a 5% gap.
  - Flag special codes.
  - Resolve overlapping years, preferring the later report.
- `test_metrics.py`: check HHI, Gini, GL, CAGR, Lawrence index and margin decomposition
  against small examples worked out by hand.

**Pipeline assertions** (run every time, fail loudly):
- Row-sum reconciliation per (flow, fy) within 0.1%.
- No duplicate keys.
- No negative values.
- All `hs2` mapped.
- The number of fiscal years per flow is 9.

**Regression checks:** Store `reports/tables/totals_check.csv`. If a refresh changes
historical totals, the diff shows in git.

**Spot checks:**
- Compare 3–5 headline numbers (annual totals, crude oil imports, smartphone exports)
  with official press releases from the Ministry of Commerce or DGCI&S.
- Record each source and date in the report appendix.

---

## 14. Conventions

- **Fiscal years:**
  - Store `fy` as `"2023-24"` and `fy_start` as 2023.
  - Display as `FY2023-24`, or `FY24` in tight chart labels. FY24 means April 2023 to
    March 2024.
- **Units:**
  - Store US$ million.
  - Charts may show US$ billion, with the unit stated on the axis.
  - Never mix ₹ and US$.
- **Codes:** Always strings, zero-padded to 8 digits, never ints.
- **Naming:**
  - snake_case for columns and functions.
  - Figures follow `<flow>_<topic>_<level>_<period>.png`.
- **Git:**
  - One commit per change, with descriptive messages.
  - Raw data is committed.
  - Processed data is regenerated (or committed as a snapshot; see §17).
- **Notebooks:** They run top to bottom on a fresh kernel. No hidden state, and no
  analysis that exists only in notebook output.

---

## 15. Risks, limitations and how we handle them

| Risk / limitation | Effect | Handling |
|---|---|---|
| Source errors (e.g. FY2023-24 export doubling) | Wrong shares, growth and rankings | Detect via total reconciliation; correct with an assertion; log it; re-check after refresh |
| HS code churn and the HS 2022 revision | Spurious entries and exits; broken HS8 series | Use HS6/HS4 for time series; stable panel; optional concordance |
| Nominal US$ only | Price, currency and volume effects mixed together | State it in every interpretation; optional quantity data (§16) |
| Only 9 annual points | Low statistical power; no ARIMA or seasonality work | Descriptive emphasis; non-parametric tests; bootstrap; honest caveats |
| Re-imports and re-exports included | Some flows aren't domestic consumption or production | Mention in limitations; can't be separated with this data |
| Suppressed or unspecified lines | Small unallocated value | Flag; keep in totals; leave out of product metrics |
| Very large single lines (crude, gold) | Concentration and totals driven by commodity prices | Report metrics with and without chapters 27 and 71 |
| TradeStat site changes | Downloader breaks | `inspect` / `assist` fallbacks; selectors kept in one config block |
| Rounding to 0.01 US$ million | Tiny lines (< 0.005) show as 0 | Treat values < 0.01 as "trace"; don't base claims on them |

---

## 16. Extensions

Optional, in rough order of value:

1. **[~] Quantity reports.** The portal offers "Quantity" at 8-digit level. Download them
   (add a `UNIT` option to the downloader), compute **unit values** (US$ per unit) and
   split value growth into price and volume parts. Units differ by product (kg, tonnes,
   numbers), so compare within a code only.
2. **[~] Country-wise and commodity-by-country reports.** These show who India buys from
   and sells to: partner concentration, China's share of electronics imports, Russia's
   share of crude, the US share of exports. They need new downloader flows; the forms
   differ.
3. **[~] Monthly data.** This would allow seasonality and higher-frequency trend and
   shock analysis.
4. **[~] Inflation adjustment.** Deflate by a US price index (e.g. US CPI or import and
   export price indices) to get rough real values. Clearly an approximation.
5. **[~] Revealed comparative advantage (Balassa RCA).** This needs world export data by
   product (e.g. UN Comtrade or CEPII BACI), comparing India's export share in product i
   with the world's.
6. **[~] Longer history.** The dropdown goes back to 2018-2019 (giving FY2017-18). Older
   years may be on other parts of the portal or in DGCI&S publications.
7. **[~] Dashboard.** An interactive explorer: pick a sector and see imports, exports,
   balance and top lines over time.

---

## 17. Open decisions

Each has a suggested default; revisit if needed.

| Decision | Options | Default |
|---|---|---|
| Processed-data format | parquet / CSV | **parquet** (types preserved, small); CSV exports for tables in `reports/` |
| Commit processed data? | yes / no | **no** (regenerate with the pipeline); revisit if collaborators lack the environment |
| Main time-series HS level | HS8 / HS6 / HS4 | **HS6**, with HS4 for long series across the HS 2022 revision |
| Treatment of special lines | include / exclude in metrics | **exclude** from product metrics, **include** in totals |
| Deflation | none / US CPI | **none** (nominal), with caveats; deflation as an extension |
| Overlap conflicts (if revisions appear) | earlier / later report | **later** report wins, with a logged diff |
| Concentration with or without crude and gold | both / one | **both**, reported side by side |
| Sector grouping | official sections only / custom groups | **both**: sections for rigour, custom groups for storytelling |

---

## 18. Milestone checklist

**M0 — Acquisition** ✅
- [x] Downloader working; 16 reports committed; manifest; defect documented.

**M1 — Clean data**
- [ ] Add `pyarrow` and `pytest` to requirements.
- [ ] `src/trade/load.py`, `clean.py`, `pipeline.py`.
- [ ] Processed tables and cleaning log.
- [ ] Load and clean tests passing.
- [ ] `01_cleaning.ipynb`.

**M2 — Enrichment**
- [ ] `hs_chapters.csv`, `hs_sections.csv`, `sector_groups.csv`.
- [ ] Stable-panel flags and coverage table.
- [ ] HS-level policy confirmed (§8.3).

**M3 — EDA**
- [ ] `02_eda_imports.ipynb`, `03_eda_exports.ipynb` (§9.1–9.6).
- [ ] Shared plot style; figures exported.

**M4 — Statistics**
- [ ] `metrics.py` with tests.
- [ ] `04_statistics.ipynb` (§10.1–10.9); tables exported.

**M5 — Balance and intra-industry trade**
- [ ] `05_balance_and_iit.ipynb`; linked-pair reference table.

**M6 — Report**
- [ ] `reports/report.md`, `README.md`.
- [ ] Spot-check headline numbers against official releases.
- [ ] Freeze `requirements.lock`.

**Stretch**
- [ ] Quantity data and unit values; country-wise data; dashboard.
