# Agensi listing: Sheets Formula Coach

DRAFT. Nothing has been submitted. The owner pastes these fields into agensi.io/dashboard/submit herself.

| | |
|---|---|
| ZIP to upload | `sheets-formula-coach\dist\sheets-formula-coach.zip` (built by `scripts\build-agensi-zip.ps1`) |
| ZIP contents | SKILL.md, references/ (6), examples/ (1), CHANGELOG.md. No README, LICENSE, evals or tests |
| Version | 1.1.0 |
| Price | **Free** |
| Brand | Codelady Abroad |

## Name (≤ 60)
```text
Sheets Formula Coach: Excel & Google Sheets, Any Locale
```

## Summary (≤ 300)
```text
Free Excel and Google Sheets formula coach. Plain words in, a working formula out, for your version and locale (; or , and Russian/Turkish function names), explained piece by piece with a traced example. Fixes errors, cleans messy data, sets up pivots, charts, conditional formatting and templates.
```

## Description
```markdown
Describe what you need in plain words, or paste a broken formula or a few rows. You get a formula that works in your app, version and language, and you learn why it works.

You get:
- the formula for Excel or Google Sheets, with `;` or `,` to match your locale
- Russian and Turkish Excel function names (e.g. ВПР / DÜŞEYARA), checked against Microsoft's pages
- a version check, with fallbacks when a function isn't in Excel 2016/2019 (e.g. XLOOKUP → INDEX/MATCH)
- a piece-by-piece explanation and a traced mini example, or test data to paste in
- fixes for #N/A, #REF!, #VALUE!, #NAME?, #SPILL! and parse errors
- data-cleaning recipes, conditional formatting and formula auditing
- pivot tables and charts step by step
- templates: budget, invoice tracker, study plan, grade book, loan payment

**Where it stops:** it can't open your file and works from what you paste. It never hides errors with IFERROR as a first fix. It gives no macros that delete files or send e-mail without explaining every line and the risk. Loan maths is not financial advice.
```

## Categories (up to 3)
1. Productivity
2. Education & Learning
3. Business Operations

## Tags (10)
excel, google-sheets, spreadsheets, formulas, data-cleaning, pivot-tables, students, beginners, multilingual, free

## Use cases (8)
1. Turn a plain-language request into an Excel or Sheets formula
2. Fix a formula that shows #N/A, #REF!, #NAME? or a parse error
3. Get the formula with semicolons and Russian or Turkish names
4. Check if a function works in Excel 2016, 2019, 2021 or 365
5. Clean messy data such as spaces, text numbers, dates, duplicates
6. Build a pivot table and chart step by step
7. Highlight overdue or duplicate rows with conditional formatting
8. Start a budget, invoice tracker, grade book or study plan

## Compatibility note
```text
Text only: no scripts, no file access, no add-ins. Works on any agent that can load a SKILL.md. Paste a few rows (with fake names) for the best answer.
```

## "See it in action" demo
Paste `examples/elvin-turkish-excel-2019-lookup.md` (request + full result, about 1,600 characters; add one more paragraph or a second short case to reach the recommended 700+ characters of output, which it already passes).

## Limitations / FAQ
- Function names marked † in the locale table are common names not re-checked on Microsoft's site this time. The skill tells users to pick the name from autocomplete.
- No evidence that Azerbaijani Excel translates function names, so the skill asks which names the user sees.
- Microsoft's old Functions Translator add-in is discontinued.

## Docs URL
Leave empty until the GitHub repo is public (then `https://github.com/codelady-ship/sheets-formula-coach` and uncomment `docs_url` in SKILL.md).

## Evidence you may quote
Blind evals: see `tests/blind-eval-results-2026-11-11.md`.
