# Functions (quick reference)

Shown with `,` separators. In `;` locales replace every argument comma with `;`.
"365" means Excel 365 or Excel 2021 and newer. "Sheets" means Google Sheets.

## Lookups

| Need | Formula | Works in |
|---|---|---|
| Find a price by product name | `=XLOOKUP(E2, A2:A100, C2:C100, "not found")` | 365, Sheets |
| Same, older Excel | `=INDEX(C2:C100, MATCH(E2, A2:A100, 0))` | all |
| Same with VLOOKUP | `=VLOOKUP(E2, A2:C100, 3, FALSE)` | all |
| Lookup with two conditions | `=XLOOKUP(1, (A2:A100=E2)*(B2:B100=F2), C2:C100)` | 365, Sheets |
| Last match instead of first | `=XLOOKUP(E2, A2:A100, C2:C100, , 0, -1)` | 365, Sheets |

Pieces of XLOOKUP: what to find, where to look, what to return, what to show if missing, match mode, search direction.
VLOOKUP traps: the lookup column must be the first column of the range; the column number breaks if a column is inserted; forgetting `FALSE` gives wrong "nearest" results.

## Conditional sums and counts

| Need | Formula |
|---|---|
| Total spent on "Food" | `=SUMIFS(C:C, B:B, "Food")` |
| Food in November 2026 | `=SUMIFS(C:C, B:B, "Food", A:A, ">="&DATE(2026,11,1), A:A, "<"&DATE(2026,12,1))` |
| How many students passed (≥ 51) | `=COUNTIFS(D2:D60, ">=51")` |
| Average score of group A | `=AVERAGEIFS(D2:D60, B2:B60, "A")` |
| Count non-empty cells | `=COUNTA(A2:A100)` |
| Count text containing "visa" | `=COUNTIFS(A2:A100, "*visa*")` |

Rule: the sum range comes first in SUMIFS, then pairs of (range, condition). All ranges must be the same size. Conditions with operators go in quotes and join cells with `&`.

## Many results (dynamic arrays)

| Need | Formula | Works in |
|---|---|---|
| All rows where status is "Unpaid" | `=FILTER(A2:D100, D2:D100="Unpaid", "none")` | 365, Sheets |
| List without duplicates | `=UNIQUE(B2:B100)` | 365, Sheets |
| Sorted by column 3, biggest first | `=SORT(A2:D100, 3, -1)` | 365, Sheets |
| Sum per category (Sheets) | `=QUERY(A1:C100, "select B, sum(C) where B is not null group by B", 1)` | Sheets |

These "spill" into the cells below and right. Keep those cells empty.

## Logic

| Need | Formula |
|---|---|
| Pass / fail | `=IF(D2>=51, "Pass", "Fail")` |
| Grades | `=IFS(D2>=91, "A", D2>=81, "B", D2>=71, "C", D2>=61, "D", D2>=51, "E", TRUE, "F")` |
| Both conditions | `=IF(AND(D2>=51, E2="yes"), "OK", "Check")` |
| Show blank instead of #N/A | `=IFNA(XLOOKUP(...), "")` |

## Text

| Need | Formula |
|---|---|
| Remove extra spaces | `=TRIM(A2)` |
| Join first and last name | `=A2&" "&B2` or `=TEXTJOIN(" ", TRUE, A2:B2)` |
| Text before "@" | `=TEXTBEFORE(A2, "@")` (365) / `=LEFT(A2, FIND("@", A2)-1)` (all) |
| Split by comma (Sheets) | `=SPLIT(A2, ",")` |
| Number stored as text → number | `=VALUE(A2)` or `=A2*1` |
| Capitalise names | `=PROPER(A2)` |

## Dates

| Need | Formula |
|---|---|
| Today | `=TODAY()` |
| Days left until deadline | `=B2-TODAY()` (format as number) |
| Same day next month | `=EDATE(A2, 1)` |
| Last day of the month | `=EOMONTH(A2, 0)` |
| Working days between | `=NETWORKDAYS(A2, B2)` |
| Age in full years | `=DATEDIF(A2, TODAY(), "Y")` |
| Month name | `=TEXT(A2, "mmmm")` (format codes can differ by locale) |

Dates are numbers underneath. If a date is left-aligned, it is probably text; fix it with Data → Text to Columns (Excel) or by retyping / DATEVALUE.

## Absolute references

`A2` moves when copied, `$A$2` never moves, `$A2` keeps the column, `A$2` keeps the row. Press F4 (Excel, Windows) to cycle. Use `$` for a fixed rate or a lookup table.
