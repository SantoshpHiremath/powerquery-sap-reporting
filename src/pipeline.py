"""End-to-end demo: generate the simulated SAP export + budget plan,
clean and merge them, run plausibility checks, and print a KPI report
summary — the posting's "Erstellung, Plausibilisierung und Visualisierung
von Kennzahlenreports" and "regelmäßige Reportings / Ad-hoc-Auswertungen"
tasks.
"""
from pathlib import Path

from sap_export_simulator import generate_sap_export
from budget_plan_generator import generate_budget_plan
from transform import load_sap_export, load_budget_plan, merge_actuals_and_budget, plausibility_checks, kpi_summary

DATA_DIR = Path(__file__).parent.parent / "data"


def run():
    sap_path = DATA_DIR / "sap_controlling_export_simulated.csv"
    budget_path = DATA_DIR / "budget_plan_2026.xlsx"

    if not sap_path.exists():
        generate_sap_export(sap_path)
    if not budget_path.exists():
        generate_budget_plan(budget_path)

    print("=== Loading & Cleaning SAP Export ===")
    actuals = load_sap_export(sap_path)
    print(f"  {len(actuals)} valid transaction rows after removing blanks and the summary row")
    print(f"  Sample: {actuals.iloc[0].to_dict()}\n")

    print("=== Loading Budget Plan (Excel) ===")
    budget = load_budget_plan(budget_path)
    print(f"  {len(budget)} budget rows across {budget['Kostenstelle'].nunique()} cost centers\n")

    print("=== Merging Actuals vs. Plan ===")
    merged = merge_actuals_and_budget(actuals, budget)
    print(f"  {len(merged)} cost-center/period combinations\n")

    print("=== KPI Summary (regelmäßiges Reporting) ===")
    kpis = kpi_summary(merged)
    for k, v in kpis.items():
        print(f"  {k}: {v}")

    print("\n=== Plausibility Check: cost centers >25% off plan (Ad-hoc-Auswertung) ===")
    flagged = plausibility_checks(merged, threshold_pct=25.0)
    print(f"  {len(flagged)} of {len(merged)} rows flagged")
    for _, row in flagged.head(10).iterrows():
        print(f"    {row['Kostenstelle']} ({row['Kostenstellenbezeichnung']}), {row['Periode']}: "
              f"Ist={row['Ist_EUR']:.2f} EUR, Plan={row['Plan_EUR']:.2f} EUR, "
              f"Abweichung={row['Abweichung_Pct']:.1f}%")


if __name__ == "__main__":
    run()
