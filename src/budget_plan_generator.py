"""Generates a synthetic Excel budget-plan workbook — the "Excel" data
source named in the posting, alongside the simulated SAP export.

Disclosure: synthetic data, not a real BR budget plan.
"""
import random
from pathlib import Path

import pandas as pd

COST_CENTERS = ["0004711", "0004722", "0004733", "0004744", "0004755"]


def generate_budget_plan(path, seed=42):
    """Writes the budget plan to .xlsx with Kostenstelle and Periode
    forced to Excel's Text format. Without this, Excel/openpyxl's
    default type inference mangles both columns on write AND on
    read-back: "0004711" becomes the number 4711 (leading zeros lost)
    and "12.2026" gets parsed as the float 12.2026 (or worse, treated as
    a date). This is a real, well-known Excel gotcha with SAP-style
    zero-padded IDs and dotted period strings — caught by
    test_budget_plan_preserves_leading_zeros_and_period_format in this
    project's test suite, not just assumed away.
    """
    # Base monthly budget per cost center, set to match the realistic
    # scale of the simulated SAP export's actual transaction volume
    # (multiple transactions per cost center per month, averaging
    # ~45-50K EUR/month per cost center) — an earlier version of this
    # generator used single-transaction-scale bases (~8-11K), which made
    # nearly every cost center look ~500-1500% over budget. That was a
    # data-modeling bug (inconsistent granularity between the two
    # synthetic sources), not a real finding, and was caught by manually
    # inspecting the KPI summary output before trusting it — see
    # test_budget_plan_scale_is_consistent_with_actuals in the test
    # suite, added as a regression guard against this exact mistake.
    rng = random.Random(seed)
    rows = []
    for cc in COST_CENTERS:
        for period in range(1, 13):
            base = {"0004711": 48000, "0004722": 46000, "0004733": 49000,
                    "0004744": 43000, "0004755": 44000}[cc]
            plan = base * (1 + rng.gauss(0, 0.08))
            rows.append({
                "Kostenstelle": cc,
                "Periode": f"{period:02d}.2026",
                "Plan_EUR": round(plan, 2),
            })

    df = pd.DataFrame(rows)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Budgetplan")
        worksheet = writer.sheets["Budgetplan"]
        # Kostenstelle = column A, Periode = column B (1-indexed with header row 1)
        kostenstelle_col = df.columns.get_loc("Kostenstelle") + 1
        periode_col = df.columns.get_loc("Periode") + 1
        for row in range(2, len(df) + 2):
            worksheet.cell(row=row, column=kostenstelle_col).number_format = "@"
            worksheet.cell(row=row, column=periode_col).number_format = "@"

    return path


if __name__ == "__main__":
    out = generate_budget_plan("data/budget_plan_2026.xlsx")
    print(f"Wrote budget plan to {out}")
