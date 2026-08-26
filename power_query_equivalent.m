// Power Query (M language) — SAP controlling export cleaning query.
//
// DISCLOSURE: This M-code was hand-written to exactly match the logic in
// src/transform.py's load_sap_export() function. It was NOT exported
// from a working Power Query session, because Power Query requires
// Windows Excel or Power BI Desktop to author and execute, and this
// project was built in a Linux environment without either. The Python
// version is what was actually run and tested; this file documents what
// the equivalent Power Query steps would be in a real Excel/Power BI
// environment, step-for-step. If Power BI Desktop is available in an
// interview or on the job, this query can be pasted into the Advanced
// Editor and adapted directly.

let
    // Step 1: Load the SAP export (semicolon-delimited, SAP's typical
    // export format) with all columns as text so amount parsing is
    // handled explicitly in a later step, not auto-detected incorrectly.
    Source = Csv.Document(
        File.Contents("sap_controlling_export_simulated.csv"),
        [Delimiter=";", Columns=8, Encoding=65001, QuoteStyle=QuoteStyle.None]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),

    // Step 2: Drop fully blank rows (SAP ALV export artifacts).
    RemoveBlankRows = Table.SelectRows(
        PromotedHeaders,
        each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))
    ),

    // Step 3: Drop the trailing "Summe" (total) footer row — it is not
    // a data row and would double-count totals if left in.
    RemoveSummaryRow = Table.SelectRows(RemoveBlankRows, each [Kostenstelle] <> "Summe"),

    // Step 4: Parse German-formatted SAP amounts, including the
    // trailing-minus convention for negative values (e.g. "1.234,56-").
    // Power Query's Locale-aware Number.FromText does not handle the
    // trailing-minus convention, so it is parsed explicitly.
    AddParsedAmount = Table.AddColumn(RemoveSummaryRow, "Betrag_EUR", each
        let
            raw = Text.Trim([Betrag]),
            isNegative = Text.EndsWith(raw, "-"),
            stripped = if isNegative then Text.Start(raw, Text.Length(raw) - 1) else raw,
            noThousands = Text.Replace(stripped, ".", ""),
            dotDecimal = Text.Replace(noThousands, ",", "."),
            asNumber = Number.FromText(dotDecimal)
        in
            if isNegative then -asNumber else asNumber,
        type number
    ),

    // Step 5: Keep Kostenstelle as text explicitly — it is an
    // identifier with meaningful leading zeros, not a number.
    SetTypes = Table.TransformColumnTypes(AddParsedAmount, {
        {"Kostenstelle", type text},
        {"Kostenstellenbezeichnung", type text},
        {"Kostenart", type text},
        {"Kostenartbezeichnung", type text},
        {"Periode", type text},
        {"Buchungsdatum", type text},
        {"Währung", type text}
    }),

    RemoveOriginalAmountColumn = Table.RemoveColumns(SetTypes, {"Betrag"})
in
    RemoveOriginalAmountColumn


// --- Second query: budget plan (Excel source) ---
// let
//     Source = Excel.Workbook(File.Contents("budget_plan_2026.xlsx"), null, true),
//     Budgetplan_Sheet = Source{[Item="Budgetplan", Kind="Sheet"]}[Data],
//     PromotedHeaders = Table.PromoteHeaders(Budgetplan_Sheet, [PromoteAllScalars=true]),
//     SetTypes = Table.TransformColumnTypes(PromotedHeaders, {
//         {"Kostenstelle", type text}, {"Periode", type text}, {"Plan_EUR", type number}
//     })
// in
//     SetTypes


// --- Third query: merge actuals + plan (Merge Queries in the Power
// Query UI, equivalent to a pandas outer merge on Kostenstelle+Periode) ---
// let
//     Source = Table.NestedJoin(
//         SAP_Actuals_Grouped, {"Kostenstelle", "Periode"},
//         Budget_Plan, {"Kostenstelle", "Periode"},
//         "BudgetMatch", JoinKind.FullOuter
//     ),
//     Expanded = Table.ExpandTableColumn(Source, "BudgetMatch", {"Plan_EUR"}),
//     AddVariance = Table.AddColumn(Expanded, "Abweichung_EUR", each ([Ist_EUR] ?? 0) - ([Plan_EUR] ?? 0))
// in
//     AddVariance
