"""Generates a simulated SAP-style controlling export.

Disclosure: this is NOT a real SAP export — there is no SAP system access
available in this environment. It is a synthetic CSV built to mirror the
real, well-known quirks of SAP table/report exports so the cleaning logic
downstream has to handle genuine SAP-shaped problems, not an idealized
CSV:

  - German number formatting (comma as decimal separator, period as
    thousands separator: "1.234,56")
  - Cost-center codes with leading zeros stored as text (e.g. "0004711")
  - A trailing summary/total row that must be excluded from analysis
  - Some blank rows (common in SAP ALV export-to-Excel/CSV output)
  - Column headers in German, matching typical SAP controlling reports
    (Kostenstelle, Kostenart, Betrag, Periode, Buchungsdatum)
  - Negative amounts marked with a trailing "-" instead of a leading
    minus sign (a well-known SAP export convention: "1.234,56-")
"""
import csv
import random
from pathlib import Path

COST_CENTERS = [
    ("0004711", "Programmdirektion Hörfunk"),
    ("0004722", "Programmdirektion Fernsehen"),
    ("0004733", "Technik & Produktion"),
    ("0004744", "IT & Digitalisierung"),
    ("0004755", "Verwaltung & Controlling"),
]

COST_TYPES = [
    ("600100", "Personalkosten"),
    ("620200", "Sachkosten"),
    ("630300", "Produktionskosten"),
    ("640400", "Lizenzkosten"),
    ("650500", "IT-Kosten"),
]


def _format_sap_amount(value):
    """Formats a number the way SAP typically exports it: German decimal
    style, trailing minus for negatives."""
    negative = value < 0
    value = abs(value)
    formatted = f"{value:,.2f}"
    # swap , and . to get German style (1,234.56 -> 1.234,56)
    formatted = formatted.replace(",", "§").replace(".", ",").replace("§", ".")
    if negative:
        formatted += "-"
    return formatted


def generate_sap_export(path, n_rows=400, seed=42):
    """Writes a simulated SAP controlling export CSV to `path`."""
    rng = random.Random(seed)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    rows = []
    for i in range(n_rows):
        cc_code, cc_name = rng.choice(COST_CENTERS)
        ct_code, ct_name = rng.choice(COST_TYPES)
        period = rng.randint(1, 12)
        day = rng.randint(1, 28)
        amount = rng.gauss(8000, 3500)
        # ~8% of rows are credits/reversals (negative amounts)
        if rng.random() < 0.08:
            amount = -abs(amount) * 0.3

        rows.append({
            "Kostenstelle": cc_code,
            "Kostenstellenbezeichnung": cc_name,
            "Kostenart": ct_code,
            "Kostenartbezeichnung": ct_name,
            "Periode": f"{period:02d}.2026",
            "Buchungsdatum": f"{day:02d}.{period:02d}.2026",
            "Betrag": _format_sap_amount(amount),
            "Währung": "EUR",
        })

    fieldnames = ["Kostenstelle", "Kostenstellenbezeichnung", "Kostenart",
                  "Kostenartbezeichnung", "Periode", "Buchungsdatum", "Betrag", "Währung"]

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for idx, row in enumerate(rows):
            writer.writerow(row)
            # sprinkle in a few blank rows, mimicking SAP ALV export artifacts
            if idx % 137 == 0 and idx > 0:
                writer.writerow({k: "" for k in fieldnames})
        # trailing summary/total row — must be excluded from analysis
        total = sum(
            float(r["Betrag"].replace(".", "").replace(",", ".").rstrip("-")) * (-1 if r["Betrag"].endswith("-") else 1)
            for r in rows
        )
        writer.writerow({
            "Kostenstelle": "Summe",
            "Kostenstellenbezeichnung": "",
            "Kostenart": "",
            "Kostenartbezeichnung": "",
            "Periode": "",
            "Buchungsdatum": "",
            "Betrag": _format_sap_amount(total),
            "Währung": "EUR",
        })

    return path


if __name__ == "__main__":
    out = generate_sap_export("data/sap_controlling_export_simulated.csv")
    print(f"Wrote simulated SAP export to {out}")
