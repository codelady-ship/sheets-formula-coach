# Data cleaning, conditional formatting, auditing and test data

Formulas use `,`; switch to `;` for comma-decimal locales.

## Cleaning recipes

| Problem | Fix |
|---|---|
| Extra or double spaces | `=TRIM(A2)`; non-breaking spaces from the web: `=TRIM(SUBSTITUTE(A2, CHAR(160), " "))` |
| Hidden characters | `=CLEAN(A2)` |
| Numbers stored as text (left-aligned, green triangle) | `=VALUE(A2)` or `=A2*1`; or select → the warning icon → Convert to Number |
| "1 250,50 AZN" as text (comma-decimal locale, so shown with `;`) | `=VALUE(SUBSTITUTE(SUBSTITUTE(A2;" AZN";"");" ";""))` |
| Dates as text "05.11.2026" | Excel: Data → Text to Columns → Date (DMY). Formula: `=DATE(RIGHT(A2,4), MID(A2,4,2), LEFT(A2,2))` |
| Split "Name Surname" | 365: `=TEXTSPLIT(A2, " ")`; Sheets: `=SPLIT(A2, " ")`; older: `=LEFT(A2, FIND(" ",A2)-1)` and `=MID(A2, FIND(" ",A2)+1, 100)` |
| Upper/lower case mix | `=PROPER(A2)`, `=UPPER(A2)`, `=LOWER(A2)` |
| Find duplicates | Helper column `=COUNTIF($A$2:$A$500, A2)>1` → TRUE = duplicate; or conditional formatting → Duplicate values |
| Remove duplicates | Excel: Data → Remove Duplicates (make a copy first). Sheets: Data → Data cleanup → Remove duplicates. Formula list: `=UNIQUE(A2:A500)` (365/Sheets) |
| Blank cells counted wrongly | `COUNTA` counts formulas that return "" too; use `=COUNTIF(A2:A100, "?*")` for real text, or `=SUMPRODUCT(--(LEN(A2:A100)>0))` |
| Blank vs zero | `=IF(A2="", "", A2*2)` keeps blanks blank |

Order: **1. make a copy of the file (or sheet)** → 2. clean with helper formulas → 3. copy → Paste Special → **Values** → 4. delete the helper column → 5. Remove Duplicates if needed.

## Conditional formatting (no formulas in cells needed)

Excel: Home → Conditional Formatting. Sheets: Format → Conditional formatting.
- Overdue dates: rule "Format only cells that contain" / "Date is before today", or a formula rule `=AND($D2<TODAY(), $G2<>"Paid")` applied to the whole table (the `$` keeps the column fixed).
- Highlight a whole row by status: select A2:H500 → formula rule `=$G2="Overdue"`.
- Top/bottom 10, data bars, colour scales: built-in presets.
- Duplicates: built-in "Duplicate values" (Excel) or `=COUNTIF($A$2:$A$500, A2)>1` (Sheets).
- Rule order matters; "Stop if true" (Excel) or put the most important rule first (Sheets).

## Auditing (find why a result is wrong)

1. Click the cell → read the formula bar.
2. Excel: Formulas → **Evaluate Formula** (step through), **Trace Precedents** (arrows to the inputs), **Error Checking**. Sheets: click the cell, hover over the error, or test pieces in a spare cell.
3. Press F9 on a selected part of the formula in the formula bar (Excel, Windows) to see its value, then Esc.
4. Check calculation mode: Formulas → Calculation Options → Automatic.
5. **Volatile functions** (OFFSET, INDIRECT, TODAY, NOW, RAND) recalculate all the time and slow big files; prefer INDEX or direct ranges. INDIRECT also breaks silently when sheets are renamed.

## LET (365, 2021, 2024, Sheets)

Names a part once and reuses it: easier to read, faster.
`=LET(total, SUMIFS(C:C, B:B, E2), IF(total=0, "none", total))`
Older Excel: repeat the part, or use a helper column.

## Sheets-only helpers

- `=ARRAYFORMULA(B2:B*C2:C)` fills a whole column with one formula (wrap with `IF(B2:B="", "", …)` to skip blanks).
- `=IMPORTRANGE("spreadsheet URL", "Sheet1!A1:D100")` pulls from another file (asks for access once; only use files you are allowed to share).
- `=GOOGLEFINANCE("CURRENCY:USDAZN")` for an approximate exchange rate (delayed; not for official accounting).
- `=QUERY(A1:D100, "select B, sum(D) group by B", 1)` for SQL-like summaries.

## Test data the user can paste

Give 5–8 fake rows in a code block, tab-separated so they paste into columns, with names clearly fake and a few traps (a blank, a text number, a trailing space). Example:

```
Date	Category	Amount	Type
01.11.2026	Food	12,50	Expense
02.11.2026	Rent 	150	Expense
05.11.2026	Stipend	400	Income
07.11.2026	Food	'8	Expense
09.11.2026			Expense
```
This is the comma-decimal version (`12,50`); in a US-style locale write `12.50`. Say which traps are inside (trailing space in "Rent ", text number `'8`, blank row) and what the formula should return.
