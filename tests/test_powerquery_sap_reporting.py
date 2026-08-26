"""Test suite for the SAP-export cleaning + budget-plan merge + KPI
reporting project."""
import sys
from pathlib import Path

import pytest
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from sap_export_simulator import generate_sap_export, _format_sap_amount
from budget_plan_generator import generate_budget_plan
from transform import (
    parse_sap_amount, load_sap_export, load_budget_plan,
    merge_actuals_and_budget, plausibility_checks, kpi_summary,
)

TMP_DIR = Path(__file__).parent / "_tmp_test_data"


@pytest.fixture(scope="module")
def data_files():
    TMP_DIR.mkdir(exist_ok=True)
    sap_path = TMP_DIR / "sap_export.csv"
    budget_path = TMP_DIR / "budget_plan.xlsx"
    generate_sap_export(sap_path, n_rows=400, seed=42)
    generate_budget_plan(budget_path, seed=42)
    yield sap_path, budget_path


# --- Amount parsing (German/SAP number format) --------------------------

def test_parse_sap_amount_positive():
    assert parse_sap_amount("1.234,56") == pytest.approx(1234.56)


def test_parse_sap_amount_negative_trailing_minus():
    assert parse_sap_amount("1.234,56-") == pytest.approx(-1234.56)


def test_parse_sap_amount_no_thousands_separator():
    assert parse_sap_amount("56,78") == pytest.approx(56.78)


def test_parse_sap_amount_large_number_with_multiple_thousand_groups():
    assert parse_sap_amount("1.234.567,89") == pytest.approx(1234567.89)


def test_parse_sap_amount_blank_returns_none():
    assert parse_sap_amount("") is None
    assert parse_sap_amount(None) is None


def test_format_and_parse_roundtrip():
    for value in [1234.56, -987.65, 0.5, -0.01, 1000000.0]:
        formatted = _format_sap_amount(value)
        parsed = parse_sap_amount(formatted)
        assert parsed == pytest.approx(value)


# --- SAP export loading/cleaning -----------------------------------------

def test_load_sap_export_drops_summary_row(data_files):
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    assert "Summe" not in df["Kostenstelle"].values


def test_load_sap_export_drops_blank_rows(data_files):
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    assert df["Kostenstelle"].apply(lambda x: str(x).strip() != "").all()


def test_load_sap_export_preserves_leading_zeros(data_files):
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    # every real Kostenstelle in this dataset is 7 digits with a leading zero
    assert all(len(str(cc)) == 7 for cc in df["Kostenstelle"].unique())
    assert any(str(cc).startswith("0") for cc in df["Kostenstelle"].unique())


def test_load_sap_export_amounts_are_numeric(data_files):
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    assert pd.api.types.is_float_dtype(df["Betrag_EUR"])
    assert df["Betrag_EUR"].notna().all()


def test_load_sap_export_row_count_matches_n_rows_minus_summary(data_files):
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    assert len(df) == 400  # n_rows generated, blanks+summary excluded


def test_load_sap_export_has_some_negative_amounts(data_files):
    # regression guard: if trailing-minus parsing breaks, all amounts
    # would silently come out positive
    sap_path, _ = data_files
    df = load_sap_export(sap_path)
    assert (df["Betrag_EUR"] < 0).any()


# --- Budget plan loading: regression guard for the leading-zero /
# period-as-float bug caught during development -----------------------

def test_budget_plan_preserves_leading_zeros_and_period_format(data_files):
    _, budget_path = data_files
    df = load_budget_plan(budget_path)
    # Regression guard: an earlier version of this generator let
    # pandas/openpyxl auto-infer types on write, which silently turned
    # "0004711" into the integer 4711 and "12.2026" into the float
    # 12.2026 -- both are wrong (Kostenstelle is an identifier, Periode
    # is a month.year label, neither is a number to compute with).
    assert all(str(cc).startswith("000") for cc in df["Kostenstelle"].unique())
    assert all(len(str(cc)) == 7 for cc in df["Kostenstelle"].unique())
    assert all(str(p).count(".") == 1 and len(str(p).split(".")[0]) == 2 for p in df["Periode"].unique())


def test_budget_plan_has_twelve_months_per_cost_center(data_files):
    _, budget_path = data_files
    df = load_budget_plan(budget_path)
    counts = df.groupby("Kostenstelle").size()
    assert (counts == 12).all()


# --- Merge + KPI: regression guard for the granularity-mismatch bug ----

def test_actuals_and_budget_are_same_order_of_magnitude(data_files):
    # Regression guard: an earlier version of the budget generator used
    # single-transaction-scale bases (~8-11K/month) against actuals that
    # aggregate ~6-7 transactions/month (~45-50K/month actual), making
    # almost every cost center look 500-1500% over budget -- a data-
    # modeling bug, not a real finding. Total actual and total plan for
    # the year should be within a reasonable band of each other (not
    # off by multiple times), even though individual months can vary.
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    kpis = kpi_summary(merged)
    ratio = kpis["total_ist_eur"] / kpis["total_plan_eur"]
    assert 0.5 < ratio < 2.0, f"actual/plan ratio {ratio} suggests a scale mismatch, not real variance"


def test_merge_covers_all_cost_centers(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    assert merged["Kostenstelle"].nunique() == 5


def test_merge_computes_variance_correctly(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    for _, row in merged.iterrows():
        assert row["Abweichung_EUR"] == pytest.approx(row["Ist_EUR"] - row["Plan_EUR"])


def test_plausibility_checks_only_flags_above_threshold(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    flagged = plausibility_checks(merged, threshold_pct=25.0)
    assert (flagged["Abweichung_Pct"].abs() > 25.0).all()


def test_plausibility_checks_sorted_by_absolute_deviation_descending(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    flagged = plausibility_checks(merged, threshold_pct=10.0)
    abs_devs = flagged["Abweichung_Pct"].abs().tolist()
    assert abs_devs == sorted(abs_devs, reverse=True)


def test_kpi_summary_totals_are_internally_consistent(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    kpis = kpi_summary(merged)
    assert kpis["total_abweichung_eur"] == pytest.approx(
        kpis["total_ist_eur"] - kpis["total_plan_eur"], abs=0.01
    )


def test_kpi_summary_cost_center_counts_are_sane(data_files):
    sap_path, budget_path = data_files
    actuals = load_sap_export(sap_path)
    budget = load_budget_plan(budget_path)
    merged = merge_actuals_and_budget(actuals, budget)
    kpis = kpi_summary(merged)
    assert kpis["cost_centers_total"] == 5
    assert 0 <= kpis["cost_centers_over_budget"] <= 5


def test_deterministic_given_seed():
    p1 = TMP_DIR / "det1.csv"
    p2 = TMP_DIR / "det2.csv"
    generate_sap_export(p1, n_rows=100, seed=7)
    generate_sap_export(p2, n_rows=100, seed=7)
    assert p1.read_text() == p2.read_text()
