# Sheets Formula Coach

> Say what you want in plain words, or paste a broken formula. Get a working Excel or Google Sheets formula for your locale, a piece-by-piece explanation, a tiny example, and fixes for #N/A, #REF!, #VALUE! and friends. Also pivots, charts and simple templates.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Claude Skill](https://img.shields.io/badge/Claude-Skill-D97757.svg)](SKILL.md)
[![Price: Free](https://img.shields.io/badge/price-free-brightgreen.svg)](#license)

A free, text-only skill by Codelady Abroad for students and beginners. It runs anywhere a plain prompt runs: Claude Code, the Claude app, other agents or a website form. It never needs your file or any code to run.

## What you get

- The formula for **your app** (Excel 365 / older Excel / Google Sheets) and **your locale**: `,` or `;` separators, translated Excel function names (e.g. Russian `ВПР`, Turkish `DÜŞEYARA`).
- A plain explanation of each piece, traced on a 3–5 row example.
- Error fixes: #N/A, #REF!, #VALUE!, #DIV/0!, #NAME?, #SPILL!, #NUM!, circular references, Sheets parse errors.
- XLOOKUP, INDEX/MATCH, VLOOKUP, SUMIFS, COUNTIFS, IF/IFS, FILTER, UNIQUE, SORT, QUERY, text and date functions.
- Click-by-click pivot tables and charts for Excel and Sheets, and which chart fits the question.
- Layout advice: one table per sheet, data validation, Tables, named ranges.
- Templates: student budget, invoice tracker, study timetable.
- Answers in the user's language.

## Safety

No macros or scripts that touch files, e-mail, Drive or sharing unless every line is explained, the risk is stated and the user is told to make a copy first. It reminds users never to enable macros in files from strangers, and to replace real personal data with fake data before pasting.

## Example request

> In Google Sheets I have names in column A and scores in column B on sheet "Results". On another sheet I type a name in A2 and want the score. My formula =VLOOKUP(A2,Results!A:B,2) gives #N/A for some names. My Sheets uses 1,5 for decimals.

## Install

- **Claude Code:** copy this folder to `~/.claude/skills/sheets-formula-coach/`.
- **Claude app:** zip the folder (with `SKILL.md` inside `sheets-formula-coach/`) and upload it under *Customize → Skills*.
- **Any other agent or a web form:** use the body of `SKILL.md` as the system prompt. `references/` holds optional tables.

Test prompts and expected points: [`evals/evals.json`](evals/evals.json).

## Sources and credits

Written from scratch. Informed by:
- Microsoft Support: Excel functions reference, error pages and PivotTable help
- Google Docs Editors Help: Google Sheets function list, pivot tables and charts
- The gap noted in Codelady Abroad's research: Anthropic's `xlsx` skill needs code, and no good text-only formula skill was found. No text was copied from other repos.

## License

[MIT](LICENSE) © 2026 Codelady Abroad
