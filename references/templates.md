# Templates (columns + formulas)

Formulas use `,`; switch to `;` for comma-decimal locales. Row 1 is the header; formulas shown for row 2 and dragged down (or as Table columns in Excel).

## Monthly student budget

Sheet **Data**:

| A Date | B Category | C Description | D Amount | E Type |
|---|---|---|---|---|
| 01.11.2026 | Rent | Dorm | 150 | Expense |
| 05.11.2026 | Stipend | Scholarship | 400 | Income |

- Category: Data validation drop-down from a list on sheet **Lists** (Rent, Food, Transport, Phone, Study, Fun, Other).
- Type: drop-down Income / Expense.

Sheet **Summary** (month start in B1, e.g. 01.11.2026):

| | Formula |
|---|---|
| Income this month | `=SUMIFS(Data!D:D, Data!E:E, "Income", Data!A:A, ">="&B1, Data!A:A, "<="&EOMONTH(B1, 0))` |
| Expenses this month | same with `"Expense"` |
| Left | `=B2-B3` |
| Per category (A6 = "Food") | `=SUMIFS(Data!D:D, Data!B:B, A6, Data!E:E, "Expense", Data!A:A, ">="&$B$1, Data!A:A, "<="&EOMONTH($B$1, 0))` |
| Share of expenses | `=IF($B$3=0, 0, B6/$B$3)` formatted as % |

Add a pie or bar chart from the category table.

## Invoice / client payment tracker (freelancers)

| A Invoice # | B Client | C Date issued | D Due date | E Amount | F Paid on | G Status | H Days overdue |
|---|---|---|---|---|---|---|---|

- D2 (due in 14 days): `=C2+14`
- G2: `=IF(F2<>"", "Paid", IF(TODAY()>D2, "Overdue", "Open"))`
- H2: `=IF(G2="Overdue", TODAY()-D2, "")`
- Total unpaid: `=SUMIFS(E:E, G:G, "<>Paid")` (the text header is ignored by SUMIFS)
- Conditional formatting: G = "Overdue" → red fill.
- Invoice number: type it; avoid formulas that renumber when rows are sorted.

## Study timetable and tracker

Sheet **Plan**:

| A Date | B Subject | C Topic | D Planned min | E Done min | F Done? | G Exam date |
|---|---|---|---|---|---|---|

- F2: checkbox (Sheets: Insert → Checkbox; Excel 365: Insert → Checkbox, or a Yes/No drop-down).
- Days to exam: `=G2-TODAY()`
- Hours studied per subject (sheet **Stats**, A2 = subject): `=SUMIFS(Plan!E:E, Plan!B:B, A2)/60`
- Progress: `=COUNTIFS(Plan!B:B, A2, Plan!F:F, TRUE)/COUNTIFS(Plan!B:B, A2)` formatted as %. In Excel with a Yes/No drop-down use `"Yes"` instead of `TRUE`.
- Weekly view: `=FILTER(Plan!A2:F, Plan!A2:A>=TODAY(), Plan!A2:A<TODAY()+7)` (Sheets; in Excel 365 use `Plan!A2:F1000` and multiply the conditions: `(Plan!A2:A1000>=TODAY())*(Plan!A2:A1000<TODAY()+7)`).

## Weekly timetable grid

Rows = time slots (08:00, 09:00…), columns = Monday…Sunday. Colour by subject with conditional formatting (Format → Conditional formatting → Text contains "Math" → blue). No formulas needed.
