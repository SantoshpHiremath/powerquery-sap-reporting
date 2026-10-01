# SAP + Excel Controlling Report: Cleaning, Merge & Plausibility Checks

A tested Python project that cleans a simulated SAP controlling export
and an Excel budget plan, merges them into a combined actual-vs-plan
KPI report, and flags implausible cost-center/period combinations. I
built this to work with **Power Query** specifically against **Excel
and SAP** data sources — not just general-purpose data cleaning.

## What this is (read before citing anywhere)

A full Power Query session wasn't practical here — it requires Windows
Excel or Power BI Desktop to author and execute — so here's what I
built instead:

1. **`power_query_equivalent.m`** — the actual Power Query M-code,
   hand-written to match the Python logic step-for-step (same cleaning
   steps, same order: promote headers, drop blank rows, drop the
   trailing summary row, parse German/SAP-formatted amounts including
   the trailing-minus convention, fix column types). This is a real,
   inspectable Power Query artifact, ready to paste into the Advanced
   Editor in Power BI Desktop. I wrote it by hand to mirror the tested
   Python logic, since I had no way to run and validate M-code myself.
2. **`src/transform.py`** — the actual tested implementation, in
   Python/pandas, doing the same transformation.

**The SAP export is simulated, not a real SAP extract.** I don't have
access to an SAP system, so `src/sap_export_simulator.py` generates a
synthetic CSV that deliberately reproduces well-known real SAP export
quirks (German number formatting with a trailing minus for negatives,
zero-padded cost-center codes, a trailing summary/total row, occasional
blank rows from ALV export) — not an idealized clean CSV — so the
cleaning logic has to handle genuine SAP-shaped problems.

**The budget plan is synthetic.**

This project demonstrates the core skill directly: designing and
testing a correct multi-step data-cleaning pipeline against genuinely
messy, realistically-formatted source data, and building the KPI
report/plausibility-check logic on top of it — backed by the real,
inspectable M-code for the Power Query half of the toolchain.

## What this models

- **`src/sap_export_simulator.py`** — synthetic SAP controlling export
  (~400 transactions) with realistic SAP export formatting quirks.
- **`src/budget_plan_generator.py`** — synthetic Excel budget-plan
  workbook, one row per cost center per period.
- **`src/transform.py`** — cleaning and merge logic: parses SAP's German/
  trailing-minus amount format, drops blank and summary rows, preserves
  zero-padded cost-center codes as text, merges actuals against plan,
  computes variance, and flags plausibility outliers above a configurable
  threshold — a classic controlling "Plausibilisierung" task.
- **`power_query_equivalent.m`** — the hand-written M-code equivalent of
  the above (see disclosure above).
- **`src/pipeline.py`** — runs the full flow end-to-end and prints a KPI
  summary plus the top plausibility-flagged cost-center/period rows.

## A note on two bugs caught during development

Both are documented here rather than silently fixed, because catching
and fixing them honestly is part of the evidence this project provides.

1. **Excel type-inference bug**: writing `Kostenstelle` ("0004711") and
   `Periode` ("12.2026") to `.xlsx` without forcing Excel's Text cell
   format let pandas/openpyxl silently mangle both on read-back — the
   cost-center code lost its leading zeros (became the integer `4711`)
   and the period string was parsed as the float `12.2026`. This is a
   real, well-known Excel/Power Query gotcha with SAP-style zero-padded
   IDs, not a hypothetical one. Fixed by explicitly setting the Text
   number format on those columns when writing, and by forcing `dtype=
   str` on read as a second line of defense. Guarded by
   `test_budget_plan_preserves_leading_zeros_and_period_format`.
2. **Data-modeling scale mismatch**: an earlier version of the budget
   generator set monthly budget bases at roughly the same scale as a
   *single* SAP transaction (~8,000-11,000 EUR), while the simulated SAP
   export aggregates several transactions per cost center per month
   (~45,000-50,000 EUR actual/month). That made nearly every cost center
   look 500-1,500% over budget — not a real finding, just inconsistent
   granularity between the two synthetic sources. Caught by manually
   inspecting the KPI summary output before trusting it (the aggregate
   numbers were an obvious red flag), fixed by rescaling the budget
   bases to match the actuals' real order of magnitude, and guarded
   going forward by `test_actuals_and_budget_are_same_order_of_magnitude`.

After the fix, total actual vs. total plan for the year differ by about
1.2% in aggregate — individual cost-center/months still show real
variance (multiple cells exceed a 25% plausibility threshold), which is
expected and honestly reported: each cost-center/month cell aggregates
only 2-12 transactions with real per-transaction variance, so month-to-
month noise at that granularity is genuinely high even when the yearly
total tracks plan closely — exactly the kind of pattern a controlling
report should surface for a human to investigate, not smooth away.

## Verification

22 automated tests (`tests/test_powerquery_sap_reporting.py`) covering:
German/SAP amount parsing (including the trailing-minus convention and a
format-then-parse round-trip), correct removal of blank and summary
rows, leading-zero preservation for cost-center codes, both regression
guards described above, merge/variance correctness, plausibility-
threshold filtering and sort order, and KPI-total internal consistency.

## Running it

```bash
pip install pandas openpyxl
python3 src/pipeline.py          # generates simulated sources, cleans, merges, prints KPI + plausibility report
python3 -m pytest tests/ -v      # runs all 22 tests
```
