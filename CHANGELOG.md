# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses [Semantic Versioning](https://semver.org/).

## [1.1.0] - 2026-11-11

Research-driven features before the Agensi listing (open-skills research, 2026-11-11).

### Added
- `references/locale-and-versions.md`: **Russian and Turkish function names checked against Microsoft Support** (names not re-checked are marked †), separators, Sheets array-literal separators, and a **version check table** with fallbacks for 2016/2019 (XLOOKUP, dynamic arrays, LET, LAMBDA, TEXTBEFORE, GROUPBY/PIVOTBY, REGEX, checkboxes).
- `references/cleaning-and-formatting.md`: **data-cleaning recipes** (spaces, non-breaking spaces, text numbers and dates, splitting, duplicates, blanks), **conditional formatting** rules, **formula auditing** (Evaluate Formula, Trace Precedents, F9, volatile functions), **LET**, Sheets-only helpers (ARRAYFORMULA, IMPORTRANGE, QUERY, GOOGLEFINANCE) and **paste-in test data**.
- Templates: **grade book** and **loan/instalment payment (PMT)**.
- Fixed answer order (formula → where → how → example → watch out → other version).
- `examples/elvin-turkish-excel-2019-lookup.md`.
- Agensi frontmatter (`version`, `summary`, `tags`, `keywords`, `use_cases`), `scripts/build-agensi-zip.ps1`, `LISTING.md`, evals 4–6 and blind test results in `tests/`.

### Fixed (checked on official pages, 2026-11-11)
- Removed the tip to use Microsoft's "Functions Translator" add-in: Microsoft says it is no longer available or supported.
- Turkish MATCH is `KAÇINCI`. Russian and Turkish COUNTIFS, XLOOKUP, EOMONTH and TRIM names were added from Microsoft's pages.
- Google Sheets: function names are English unless display-language names are switched on ("Always use English function names"). Previously the skill said "accepts English names in all locales".
- Excel separators follow the Windows list separator.
- Azerbaijani Excel is a partial language pack with no evidence of translated function names, so the skill asks which names the user sees.
- Version facts: XLOOKUP is not in Excel 2016 or 2019; GROUPBY and the REGEX functions are 365 only.

### Fixed (blind evals, 2026-11-11)
- Google Sheets FILTER has no if_empty argument: wrap in IFNA (fixed in functions.md and SKILL.md).
- Wrong separator in Excel shows a "problem with this formula" dialog, not #NAME?.
- #SPILL! applies to Excel 2021+ too; Sheets shows #REF! for the same problem.
- IFNA preferred for lookups, offered only after the fix; IFNA added to the RU/TR name table.
- Budget template uses `"<"&EOMONTH(...)+1`; cleaning order starts with a copy; Turkish example shows the formula without EĞERHATA first.
- #N/A answers always list every cause (incl. number vs text) in a "Watch out" line; #NAME? no longer lists the separator as a cause.

## [1.0.0] - 2026-11-10

### Added
- First release: plain words to Excel/Sheets formulas with locale separators, piece-by-piece explanations and traced examples, error fixes, pivots and charts, layout advice, templates (budget, invoice tracker, study timetable), safe macro rules. MIT.
