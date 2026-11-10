# Errors: causes and fixes

Check in this order: look at the formula bar, check the separator, check each range, then test pieces of the formula in a spare cell.

## #N/A (not available)

| Cause | Fix |
|---|---|
| The value really isn't in the list | Correct; show a message with `IFNA(..., "not found")` |
| Extra spaces ("Ali " vs "Ali") | `TRIM` both sides, or clean the column once |
| Number vs text ("101" vs 101) | Convert: `VALUE()`, `*1`, or Data → Text to Columns |
| VLOOKUP without exact match | Add `FALSE` / `0` as the last argument |
| Range doesn't include the row | Make the range bigger or use a Table / whole column |
| Lookup column isn't first in VLOOKUP | Use XLOOKUP or INDEX/MATCH |

## #REF! (bad reference)

| Cause | Fix |
|---|---|
| A referenced row, column or sheet was deleted | Undo, or rewrite the reference |
| VLOOKUP column number bigger than the range | Count columns again, or use XLOOKUP |
| Copied formula moves off the sheet edge | Use `$` to lock references |
| Link to another file that moved | Update the link or paste values |

## #VALUE! (wrong type)

| Cause | Fix |
|---|---|
| Text in maths ("100 AZN" + 5) | Keep only numbers in the cell; currency via format |
| Date stored as text | Convert to a real date |
| Ranges of different sizes in SUMIFS/FILTER | Make all ranges the same rows |
| Hidden space or non-printing character | `TRIM(CLEAN(A2))` |

## #DIV/0!

Dividing by zero or an empty cell. Fix: `=IF(B2=0, "", A2/B2)`. Check why the divisor is empty first.

## #NAME?

| Cause | Fix |
|---|---|
| Typo in the function name | Retype; use the suggestion list |
| Text without quotes (`=IF(A2=Yes, ...)`) | `"Yes"` |
| Function not in this version (XLOOKUP in Excel 2016) | Use INDEX/MATCH |
| English name in a translated Excel, or the reverse | Use the local name |
| Named range doesn't exist | Create it or fix the name |

## #SPILL! (Excel 365)

The result needs empty cells below or to the right. Clear them, or move the formula. Tables (Ctrl+T) can't hold spilling formulas; put them outside the table.

## #NUM!

Impossible number (e.g. square root of a negative, a date before 1900, an iteration that never finishes). Check the inputs.

## Circular reference

The formula refers to its own cell, directly or through other cells. Excel shows a warning and the cell in Formulas → Error Checking → Circular References. Move the total outside the range it sums.

## Google Sheets: "Formula parse error" / #ERROR!

Usually the wrong separator for the locale (`,` vs `;`), an unmatched bracket or quote, or a copied curly quote (“ ”) instead of straight quotes (" ").

## Wrong result, no error

- Calculation set to Manual (Excel: Formulas → Calculation Options → Automatic).
- Numbers stored as text are ignored by SUM.
- A relative reference moved when copied; add `$`.
- Hidden or filtered rows: SUM counts them; `SUBTOTAL(109, ...)` ignores hidden rows.
