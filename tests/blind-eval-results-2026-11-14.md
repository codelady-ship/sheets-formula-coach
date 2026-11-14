# sheets-formula-coach: blind eval results (2026-11-14)

- **Result: 2/2 new cases pass (9/9 expectations)**
- **Skill version:** 1.1.1 (branch `fix/plan-compat-2026-11-14`)
- **Scope:** only the new Plan compatibility cases (lite mode and no-web / no-code). Older cases were not re-run.
- **Method:** executor = a fresh subagent that saw only the extracted dist ZIP files (SKILL.md, references/, examples/ where present, CHANGELOG.md) and, per case, the user prompt plus the stated constraint. It did not see the expectations. Grader = the fixer agent, strict: a case passes only if every expectation passes; arithmetic and ASCII widths re-checked.

## Eval 7 - PASS (5/5)

| # | Expectation | Result | Evidence |
|---|---|---|---|
| 1 | Formula in a code block that works in Turkish Excel 2019: Turkish function names (e.g. DÜŞEYARA or İNDİS/KAÇINCI) and ; separators, no ÇAPRAZARA/XLOOKUP | PASS | `=İNDİS(Notlar!$C$2:$C$500;KAÇINCI(B2;Notlar!$A$2:$A$500;0))` in a code block; says Excel 2019 has no ÇAPRAZARA |
| 2 | Gives a one-line explanation and a short traced example | PASS | One-line "How" and a trace 101/102/103 → 42 |
| 3 | Does not write the conditional-formatting steps in full now; offers them next with a one-line 'more?'-style offer | PASS | Highlighting left for the next message: "More?" |
| 4 | Cell references follow the user's layout or state the assumption | PASS | States the assumption (ID in B2, rows 2–500) |
| 5 | Short | PASS | About 115 words |

## Eval 8 (setup: no web search) - PASS (4/4)

| # | Expectation | Result | Evidence |
|---|---|---|---|
| 1 | Does not claim to have checked Google Sheets Help in this chat | PASS | "I can't search the web in this chat" |
| 2 | Does not assert that GROUPBY exists in Google Sheets; says it is an Excel 365 function per the references, and to check Google Sheets Help for Sheets | PASS | "GROUPBY in Google Sheets is not in my checked list, so I won't say it exists (check Google Sheets Help ...)"; GROUPBY = Excel 365 |
| 3 | Gives a working Sheets formula it is sure of (QUERY, or UNIQUE + SUMIF, or a pivot table) in a code block with a piece-by-piece explanation and a short trace | PASS | QUERY with group by in a code block, piece-by-piece explanation, traced 3-row example (Baku 130, Ganja 50), semicolon-locale variant |
| 4 | Answers in English, short | PASS | English; about 270 words (on the long side, still short) |

## Answers (verbatim)

### Eval 7

````text
Excel 2019 has no XLOOKUP (ÇAPRAZARA), so use İNDİS + KAÇINCI. Turkish Excel uses `;`.

Assumed: the student ID you search for is in **B2** of your sheet; data in 'Notlar' rows 2–500.

```
=İNDİS(Notlar!$C$2:$C$500;KAÇINCI(B2;Notlar!$A$2:$A$500;0))
```

- **How:** `KAÇINCI` (MATCH) finds the row of the ID in column A (`0` = exact match); `İNDİS` (INDEX) returns the grade from column C in that row.
- **Trace:** Notlar A = 101, 102, 103 · C = 75, 42, 88 → B2 = 102 → KAÇINCI = 2 → **42**.
- #YOK (#N/A)? The ID is missing, has extra spaces, or is a number in one sheet and text in the other.

Highlighting grades below 50 comes in the next message. More?
````

### Eval 8

````text
I can't search the web in this chat, and GROUPBY in Google Sheets is not in my checked list, so I won't say it exists (check Google Sheets Help: support.google.com/docs/table/25273). GROUPBY is an **Excel 365** function. In Sheets, this works for sure:

```
=QUERY(A2:B200, "select A, sum(B) where A is not null group by A label sum(B) 'Total sales'", 0)
```

- **Where:** an empty cell, e.g. D1. It fills two columns by itself (no dragging). Keep the cells below and to the right empty.
- **How it works:**
  - `A2:B200` = your regions and sales.
  - `select A, sum(B)` = show each region and its total sales.
  - `where A is not null` = skip empty rows.
  - `group by A` = one row per region (sorted A→Z).
  - `label sum(B) 'Total sales'` = the header of the totals column.
  - `0` = your range has no header row.
- **Example:**

| A Region | B Sales |
|---|---|
| Baku | 100 |
| Ganja | 50 |
| Baku | 30 |

Result: a header row (`Total sales` above the totals), then **Baku 130**, **Ganja 50**.

- **Watch out:**
  - Sales stored as text (e.g. "100 AZN") are not summed — keep only numbers in B.
  - "Baku" and "Baku " (extra space) become two groups.
  - QUERY needs one data type per column.
- **Other locale / app:**
  - Sheets with `1,5` locale: `=QUERY(A2:B200; "select A, sum(B) where A is not null group by A label sum(B) 'Total sales'"; 0)`
  - Excel 365: `=GROUPBY(A2:A200,B2:B200,SUM)` · older Excel: pivot table (Rows = Region, Values = Sum of Sales).

Explain more?
````

