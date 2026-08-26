"""Data transformation logic for the SAP controlling export and the Excel
budget-plan source, feeding a combined KPI report.

Disclosure: Power Query itself (the "Get & Transform" engine in Excel/
Power BI) requires Windows Excel or Power BI Desktop to author and run —
neither is available in this Linux environment. What's built here is
functionally identical transformation logic in Python/pandas, applying
the exact same steps in the exact same order as the M-code query
alongside it (`power_query_equivalent.m` in this folder), so the M-code
is a real, inspectable Power Query artifact and the Python version is
what's actually tested and run in this environment. This distinction is
disclosed everywhere this project is referenced — the M-code was written
by hand to match this logic, not exported from a working Power Query
session.
"""
import re

import pandas as pd


def parse_sap_amount(raw):
    """Parses a German-formatted SAP amount string like '1.234,56' or
    '1.234,56-' (trailing minus for negatives) into a float.
    """
    if raw is None or str(raw).strip() == "":
        return None
    s = str(raw).strip()
    negative = s.endswith("-")
    if negative:
        s = s[:-1]
    s = s.replace(".", "").replace(",", ".")
    try:
        value = float(s)
    except ValueError:
        return None
    return -value if negative else value


def load_sap_export(path):
    """Loads and cleans the simulated SAP controlling export:
      1. Read semicolon-delimited CSV (SAP's typical export delimiter).
      2. Drop fully blank rows (SAP ALV export artifacts).
      3. Drop the trailing 'Summe' (total) row — it's a footer, not data.
      4. Parse German-formatted amounts (incl. trailing-minus negatives)
         into proper floats.
      5. Keep cost-center codes as zero-padded strings (they're
         identifiers, not numbers — leading zeros must be preserved).
    """
    df = pd.read_csv(path, sep=";", dtype=str, keep_default_na=False)

    # step 2: drop fully blank rows
    df = df[~(df.apply(lambda row: all(v.strip() == "" for v in row), axis=1))]

    # step 3: drop the trailing summary row
    df = df[df["Kostenstelle"] != "Summe"]

    # step 4: parse amounts
    df["Betrag_EUR"] = df["Betrag"].apply(parse_sap_amount)

    # step 5: cost-center codes stay as strings (leading zeros preserved)
    df["Kostenstelle"] = df["Kostenstelle"].astype(str)

    df = df.reset_index(drop=True)
    return df[["Kostenstelle", "Kostenstellenbezeichnung", "Kostenart",
               "Kostenartbezeichnung", "Periode", "Buchungsdatum", "Betrag_EUR", "Währung"]]


def load_budget_plan(path):
    """Loads the Excel budget-plan source (one row per cost center per
    period with a planned budget amount) — the second data source the
    posting names ("Excel, SAP und weitere Datenquellen").

    Kostenstelle and Periode are forced to str on read. Even with the
    source workbook's cells formatted as Text, pandas/openpyxl can still
    infer numeric dtypes on read in some configurations — being explicit
    here is the same defensive habit as the SAP CSV parsing (never trust
    a spreadsheet tool's automatic type inference for identifier-like
    columns), and it's covered by a regression test.
    """
    df = pd.read_excel(path, dtype={"Kostenstelle": str, "Periode": str})
    return df


def merge_actuals_and_budget(actuals_df, budget_df):
    """Merges actual spend (from the SAP export) with planned budget (from
    Excel) by cost center and period, computing variance — this is the
    "Kennzahlenreport" the posting asks for: actual vs. plan by cost
    center, ready to feed a Power BI visual.
    """
    actual_by_period = (
        actuals_df.groupby(["Kostenstelle", "Kostenstellenbezeichnung", "Periode"], as_index=False)
        .agg(Ist_EUR=("Betrag_EUR", "sum"))
    )

    merged = actual_by_period.merge(
        budget_df, on=["Kostenstelle", "Periode"], how="outer", indicator=True
    )

    merged["Ist_EUR"] = merged["Ist_EUR"].fillna(0.0)
    merged["Plan_EUR"] = merged["Plan_EUR"].fillna(0.0)
    merged["Abweichung_EUR"] = merged["Ist_EUR"] - merged["Plan_EUR"]
    merged["Abweichung_Pct"] = merged.apply(
        lambda r: (r["Abweichung_EUR"] / r["Plan_EUR"] * 100) if r["Plan_EUR"] != 0 else None,
        axis=1
    )

    return merged


def plausibility_checks(merged_df, threshold_pct=25.0):
    """Flags cost-center/period combinations where actual spend deviates
    from plan by more than `threshold_pct` — the posting's "Plausibili-
    sierung" (sanity-checking) task. Returns the flagged subset, sorted
    by absolute deviation, largest first.
    """
    flagged = merged_df[
        merged_df["Abweichung_Pct"].notna()
        & (merged_df["Abweichung_Pct"].abs() > threshold_pct)
    ].copy()
    flagged = flagged.sort_values("Abweichung_Pct", key=lambda s: s.abs(), ascending=False)
    return flagged


def kpi_summary(merged_df):
    """Rolls up KPIs the way a monthly controlling report would: total
    actual, total plan, total variance, and count of cost centers over
    budget — the recurring ("regelmäßige") reporting output.
    """
    total_ist = merged_df["Ist_EUR"].sum()
    total_plan = merged_df["Plan_EUR"].sum()
    over_budget = merged_df[merged_df["Abweichung_EUR"] > 0]["Kostenstelle"].nunique()

    return {
        "total_ist_eur": round(total_ist, 2),
        "total_plan_eur": round(total_plan, 2),
        "total_abweichung_eur": round(total_ist - total_plan, 2),
        "total_abweichung_pct": round((total_ist - total_plan) / total_plan * 100, 2) if total_plan else None,
        "cost_centers_over_budget": int(over_budget),
        "cost_centers_total": int(merged_df["Kostenstelle"].nunique()),
    }
