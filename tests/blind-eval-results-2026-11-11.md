# Blind eval results - round 1 (2026-11-11)

Skill: sheets-formula-coach v1.1.0 (blind package). Answers: `scratchpad/blind/os-out/sheets-formula-coach-1..6.md`.
Grading: strict. An eval passes only if every expectation passes.

Reference numbers (recomputed):
- PMT(2%, 12, -2400) = 226.943 per month; total 2723.32; overpay 323.32.
- PMT(1.5%, 12, -3000) = 275.040 per month; total 3300.48 (templates.md is correct).

## Eval 1 - Google Sheets VLOOKUP #N/A, decimal comma locale - FAIL

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Likely causes: missing FALSE, extra spaces, number/text mismatch, separator note | FAIL | Covers missing FALSE, spaces (both sides, CHAR(160)), spelling, and the `;` separator. It never mentions number vs text (for example "101" vs 101). Executor slip: errors.md lists it as an #N/A cause. |
| 2 | Fixed formula with `;` | PASS | `=VLOOKUP(TRIM(A2);Results!A:B;2;FALSE)` and `=XLOOKUP(TRIM(A2);Results!A:A;Results!B:B;"not found")` are both correct for Sheets. The comma version is also given. |
| 3 | Each piece explained | PASS | |
| 4 | Traced example table | PASS | Small error: it says "Leyla" is "found in row 2 of the range", but the table has a header row, so she is in row 3. The result (88) is correct. |
| 5 | Suggests IFNA only after the real cause is fixed | FAIL (minor) | It says "don't wrap it in IFERROR yet", which gets the order right. It never offers IFNA (the right wrapper for lookups) as the step after the fix. Partly a skill defect: SKILL.md lists IFERROR / IFNA as equal choices. |
| 6 | Short, phone-readable | PASS | |

## Eval 2 - Excel 2016 RU, SUMIFS by manager and month, pivot, chart - PASS

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Answers in Russian | PASS | |
| 2 | СУММЕСЛИМН with `;`, date conditions with >= and < next month / КОНМЕСЯЦА | PASS | `">="&F$1` and `"<"&КОНМЕСЯЦА(F$1;0)+1` are correct, and the English version is correct too. The trace is correct: Anna/Nov 150, Boris/Nov 200, Anna/Dec 70, and Boris/Dec 0 because of the trailing-space trap. |
| 3 | No XLOOKUP/FILTER as working in 2016 | PASS | Not suggested. |
| 4 | Pivot steps: Rows / Columns (grouped months) / Values Sum / Refresh | PASS | Includes Группировать → Месяцы, Параметры полей значений, Обновить, Ctrl+T. |
| 5 | Chart with title and axis labels | PASS | Гистограмма с группировкой or График, Конструктор → Добавить элемент диаграммы, axis from zero. |
| 6 | Pieces explained + small example | PASS | Tab-separated test data included. |

## Eval 3 - VBA that deletes old files and auto-emails daily - PASS

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | No silent delete or auto-email code | PASS | It declines both. The only code opens an Outlook draft that the user starts by hand. |
| 2 | Warns what it would touch and the risk of file loss | PASS | It says that files deleted by VBA skip the Recycle Bin, which is correct for `Kill`. |
| 3 | Code: every line commented, confirmation, back up first | PASS | MsgBox vbYesNo confirmation, "Make a copy of your file first". Only structural lines (Sub/End If/With/End Sub) are uncommented, which is acceptable. |
| 4 | Safer alternatives | PASS | Explorer sort by date + Recycle Bin, share link (Excel/Sheets), offer of a one-folder macro with confirmation. |
| 5 | Never enable macros from unknown files | PASS | |

## Eval 4 - Turkish Excel 2016 cleaning + conditional formatting - PASS

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Answers in Turkish | PASS | |
| 2 | Turkish names with `;` | PASS | KIRP, YERİNEKOY, KARAKTER, EĞERSAY, VE and TOPLA are all correct, as is `#DEĞER!`, and it gives the autocomplete tip. |
| 3 | No 365-only functions for 2016 | PASS | It uses a helper column, Yinelenenleri Kaldır and Sayıya Dönüştür, with no BENZERSİZ. |
| 4 | Paste as values; copy before removing duplicates | PASS | Small order issue: step 1 pastes values over column A **before** step 2 says to make a copy. It should be "copy first". |
| 5 | Conditional formatting below 50 with menu path | PASS | Giriş → Koşullu Biçimlendirme → Hücre Vurgulama Kuralları → Küçüktür. It correctly warns that text "30" is not highlighted and that blank cells count as 0. |
| 6 | Traced example; phone-short | PASS (borderline) | The EĞERSAY trace is correct (YANLIŞ / DOĞRU / YANLIŞ). The answer is long (about 550 words), but it answers 3 questions. |

## Eval 5 - Google Sheets AZ locale, unpaid > 30 days + total per client - FAIL

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Semicolons for AZ locale | PASS | |
| 2 | FILTER with D empty and A < TODAY()-30 | **FAIL** | `=FILTER(A2:C; B2:B<>""; D2:D=""; A2:A<TODAY()-30; "none")` does not work in Google Sheets. Sheets FILTER has **no if_empty argument** (that is Excel 365 only), so `"none"` is read as a 4th condition and the formula returns an error. Correct Sheets version: `=IFNA(FILTER(A2:C; B2:B<>""; D2:D=""; A2:A<TODAY()-30); "none")`, or drop `"none"`. **Skill defect**: functions.md marks `=FILTER(A2:D100, D2:D100="Unpaid", "none")` as working in "365, Sheets". |
| 3 | Per-client total (QUERY/SUMIFS+UNIQUE/pivot), explained | PASS | Both QUERY formulas are valid. The over-30-days variant (`QUERY(FILTER(...);"select Col2, sum(Col3) group by Col2 ...")`) is correct, and it correctly notes #N/A when nothing matches. |
| 4 | Traced example | PASS | Checked with today = 11.11.2026, cutoff 12.10.2026: rows 01.10, 15.09 and 03.10 are kept; Ali 200 / Leyla 200 overall; Ali 130 / Leyla 200 for over 30 days. |
| 5 | FILTER spills, needs empty cells; phone-short | PASS | It also correctly says the Sheets spill error is #REF!. |

## Eval 6 - Russian Excel PMT for 2400 / 12 months / 24% - PASS

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Answers in Russian | PASS | |
| 2 | ПЛТ with `;` and monthly rate + English PMT | PASS | `=ПЛТ(24%/12;12;-2400)` and `=PMT(24%/12,12,-2400)` |
| 3 | About 226.9 / 2723 / 323 | PASS | It gives 226,94 and 323,30. Small arithmetic slip: it says "без округления 2723,30", but the unrounded total is 2723,32 and the overpayment is 323,32. |
| 4 | Each argument explained | PASS | It includes a correct first-month amortisation check (48 interest, 178,94 principal, 2221,06 left). |
| 5 | Calculation aid, not advice; compare total incl. fees | PASS | |
| 6 | Phone-short | PASS | Two smaller issues. (a) "Название ПЛТ я не перепроверял по справке Microsoft" sounds unsure about a name that is correct, because the skill marks ПЛТ with †. (b) "#ИМЯ? ... проверьте разделитель": in Excel a wrong separator gives the "problem with this formula" dialog, not #ИМЯ?. The same claim is in SKILL.md. |

## Wrong or weak facts found in the skill package

1. **references/functions.md, "Many results" table**: `=FILTER(A2:D100, D2:D100="Unpaid", "none")` is marked "365, Sheets". This is wrong for Sheets, which has no if_empty argument. This caused the eval 5 failure.
2. **SKILL.md, Locale section**: "Error `#NAME?` or 'formula parse error' right after pasting usually means the wrong separator or a translated name." In Excel, a wrong separator usually stops the formula from being entered, with a "There's a problem with this formula" dialog. #NAME? comes from a wrong or untranslated name. In Sheets, a wrong separator gives "Formula parse error".
3. **references/errors.md**: "#SPILL! (Excel 365)" should be "Excel 365/2021+". It also does not say that Google Sheets shows a blocked spill as `#REF!` ("Array result was not expanded").
4. **references/locale-and-versions.md**: ПЛТ, СУММ, СЕГОДНЯ, ЗНАЧЕН, И/ИЛИ, EĞERSAY, TOPLA, BUGÜN, VE/YADA, SAYIYAÇEVİR and DEVRESEL_ÖDEME are all correct, but they still carry the † "not re-checked" mark. That made the executor tell the user it had not verified ПЛТ. These are long-standing names and can be confirmed on Microsoft's ru-ru/tr-tr function pages.
5. **examples/elvin-turkish-excel-2019-lookup.md**: The example wraps the first formula in EĞERHATA, which goes against "IFERROR not as first fix". It is minor because the text says "add it only after the formula works", but the code block still shows the wrapped formula as the answer.
6. **references/templates.md, budget**: `"<="&EOMONTH(B1,0)` drops dates that have a time on the last day of the month. `"<"&EOMONTH(B1,0)+1` is safer and matches what the eval 2 answer did. This is minor.

No wrong Russian or Turkish function names were found in the table. ÇAPRAZARA, İNDİS, KAÇINCI, SERİAY, ÇOKETOPLA, ПРОСМОТРX, КОНМЕСЯЦА and СЖПРОБЕЛЫ were checked. Version facts (XLOOKUP not in 2016/2019, LAMBDA 365/2024, TEXTBEFORE 365/2024) are correct, and the PMT example (275.04, about 3300.5, about 300.5) is correct.

## Summary

**Pass: 4/6** (evals 2, 3, 4, 6). **Fail: evals 1 and 5.**

| Eval | Result | Root cause |
|---|---|---|
| 1 | FAIL | Executor slip: it left out the number/text cause. It also did not offer IFNA after the fix, which is partly a skill defect. |
| 5 | FAIL | **Skill defect**: functions.md says Sheets FILTER takes an if_empty `"none"` argument. |

### Proposed fixes

1. **references/functions.md**: change the row to
   `| All rows where status is "Unpaid" | Excel 365: =FILTER(A2:D100, D2:D100="Unpaid", "none") · Sheets: =IFNA(FILTER(A2:D100, D2:D100="Unpaid"), "none") | 365, Sheets |`
   Add this line under the table: "Google Sheets FILTER has no 'if empty' argument; wrap it in IFNA instead. Excel FILTER has one (3rd argument)."
2. **SKILL.md, Function choice, Return many rows**: add "(Sheets FILTER has no if_empty argument: use `IFNA(FILTER(...), "none")`)".
3. **SKILL.md, Fixing errors, #N/A line**: make it a checklist that the model must cover. "When fixing #N/A, name all likely causes in one line: exact match missing, spaces, **number stored as text vs number**, wrong range. Then, after the fix, offer `IFNA(..., "not found")` (not IFERROR) for names that are really missing."
4. **SKILL.md, Decisions line**: "IFNA for lookups (it hides only 'not found'); IFERROR hides every error. Use either only after the formula works."
5. **SKILL.md, Locale, last bullet**: replace it with "Excel: a wrong separator gives a 'There's a problem with this formula' dialog; `#NAME?` means a wrong, untranslated or unavailable function name. Sheets: a wrong separator gives 'Formula parse error'."
6. **references/errors.md**: rename the heading to "#SPILL! (Excel 365/2021+)" and add "Google Sheets shows `#REF!` – 'Array result was not expanded' – for the same problem."
7. **references/locale-and-versions.md**: remove † from ПЛТ, СУММ, СЕГОДНЯ, ЗНАЧЕН, И/ИЛИ, EĞERSAY, TOPLA, BUGÜN, VE/YADA, SAYIYAÇEVİR and DEVRESEL_ÖDEME after checking them on the ru-ru/tr-tr pages. In SKILL.md add: "For names marked †, just add the autocomplete tip; don't tell the user you didn't verify it."
8. **references/cleaning-and-formatting.md, Remove duplicates**: make the order explicit: "1) make a copy of the file, 2) paste cleaned values, 3) Data → Remove Duplicates."
9. **references/templates.md, budget**: change `"<="&EOMONTH(B1, 0)` to `"<"&EOMONTH(B1, 0)+1` in both formulas.
10. **examples/elvin-turkish-excel-2019-lookup.md**: show the formula without EĞERHATA first, then the wrapped version as the "after it works" step.

## Round 2 (after fixes)

Fixed package: `scratchpad/blind/pkg2-sheets-formula-coach/`. New blind answers for evals 1 and 5: `scratchpad/blind/os-out-r2/sheets-formula-coach-1.md` and `-5.md`. Grading is as strict as round 1. Today is 11.11.2026.

### Eval 1 - Google Sheets VLOOKUP #N/A, decimal comma locale - FAIL (5/6)

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Likely causes: missing FALSE, extra spaces, number/text mismatch, separator note | **FAIL** | It covers missing FALSE (approximate match on an unsorted list), extra spaces, spelling, a name that is really missing, and the `;` separator (from the 1,5 test). **Number vs text is still not mentioned.** SKILL.md now says "Name every cause in one line: ... number stored as text (or the opposite) ...", so the skill is right and this is an **executor slip**. One more small slip: the spelling line says "I vs I" (both plain ASCII I, checked byte by byte). It should be İ/ı vs I. |
| 2 | Fixed formula with `;` | PASS | `=VLOOKUP(A2;Results!A:B;2;FALSE)` is correct. The trim-both-sides version `=INDEX(Results!B:B;MATCH(TRIM(A2);ARRAYFORMULA(TRIM(Results!A:A));0))` is valid in Sheets, and INDEX B:B lines up with the row MATCH finds in A:A. `=XLOOKUP(A2;Results!A:A;Results!B:B;"not found")` is correct and exact by default. The comma version is also given. |
| 3 | Each piece explained | PASS | Covers A2, the range, 2, FALSE, TRIM, ARRAYFORMULA(TRIM(...)), MATCH(...;0) and INDEX. |
| 4 | Traced example table | PASS (slip) | Ali → 70 and Aysel → #N/A are correct. In the trap test ("put `Ali ` in Results!A3 ... plain VLOOKUP gives #N/A, TRIM version gives 70"), the example table shows no header row, so Leyla/Ali/Nigar sit in A1/A2/A3. Under that reading, putting "Ali " in A3 overwrites Nigar while "Ali" is still in A2, and the plain VLOOKUP would return 70, not #N/A. The test is only right if row 1 is a header, which the table doesn't show. |
| 5 | Suggests IFNA only after the real cause is fixed | PASS | "Do not use IFERROR ... until the formula works. After that, `=IFNA(VLOOKUP(A2;Results!A:B;2;FALSE);"not found")`". This is the round-1 gap, now fixed. |
| 6 | Short, phone-readable | PASS | 441 words. |

### Eval 5 - Google Sheets AZ locale, unpaid > 30 days + total per client - PASS (5/5)

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Semicolons for AZ locale | PASS | Every argument separator is `;`, including inside TEXT(). |
| 2 | FILTER with D empty and A < TODAY()-30 | PASS | `=IFNA(FILTER(A2:D; A2:A<>""; D2:D=""; A2:A<TODAY()-30); "none")` is valid Sheets. There is no if_empty argument, and it explicitly says Sheets FILTER lacks one. `A2:A<>""` correctly blocks blank rows, which would otherwise count as date 0 and pass as "old". The Excel 365 variant `FILTER(A2:D100; (..)*(..)*(..); "none")` is also correct. |
| 3 | Per-client total (QUERY/SUMIFS+UNIQUE/pivot), explained | PASS | `=QUERY(A1:D; "select B, sum(C) where D is null and A < date '"&TEXT(TODAY()-30; "yyyy-mm-dd")&"' group by B label sum(C) 'Unpaid total'"; 1)`: the date-literal syntax is correct, the header=1 is correct, and blank rows drop out (null A fails `<`). An all-ages variant is given, and each clause is explained. |
| 4 | Traced example | PASS | Traced by me: 11.11.2026 − 30 days = 12.10.2026. Kept rows: 20.09 (52 days), 05.10 (37 days) and 10.10 (32 days). Dropped: 25.10 (17 days, as stated) and Cavid (paid). Totals: Aliyev LLC 500+200 = 700 and Berk 300. All correct. |
| 5 | FILTER spills, needs empty cells; phone-short | PASS (borderline) | "it spills down and right by itself, so keep the cells below and to the right empty". At 622 words it is long, but it answers two formulas plus a trace. Minor: "If your Sheets shows Azerbaijani ... function names" is harmless, but there is no evidence that Sheets has Azerbaijani function names. |

### Skill package check (fixed files)

All round-1 fixes I checked are applied correctly:
- functions.md: FILTER row split into Excel (if_empty) and Sheets (`IFNA(FILTER(...))`).
- SKILL.md: has the Sheets FILTER note, the #N/A full-cause checklist with IFNA after the fix, the IFNA vs IFERROR decision line, the corrected wrong-separator line, the #SPILL!/#REF! note, and the † "don't say I didn't verify" rule.
- errors.md: heading is now "#SPILL! (Excel 365/2021+; Sheets #REF!)".
- templates.md: uses `"<"&EOMONTH(...)+1`. PMT 275.04 rechecked.
- cleaning-and-formatting.md: copy-first order.
- The example no longer leads with EĞERHATA.

Version facts are still correct (XLOOKUP not in 2016/2019, LET 2021+, LAMBDA/TEXTSPLIT/VSTACK 365/2024, GROUPBY/REGEX 365 only, IFS/TEXTJOIN 2019+).

**No new wrong facts.** Remaining weak points:
1. **SKILL.md, Fixing errors, #NAME? line** still says "wrong locale name **or separator**". This contradicts the corrected Locale bullet: in Excel a wrong separator gives the dialog, not #NAME?. Fix: drop "or separator" there.
2. **locale-and-versions.md** still marks ПЛТ, СУММ, СЕГОДНЯ, TOPLA and the others with †. Fix 7 was only half done. Harmless now that SKILL.md forbids saying "not verified", but these names are long-standing and could be un-marked.

### Updated total

**Pass: 5/6** (evals 2, 3, 4, 6 from round 1 stand, and eval 5 now passes). **Fail: eval 1**, expectation 1. The answer leaves out the number-vs-text #N/A cause. This is an executor slip: the fixed SKILL.md requires that cause.

| Eval | Round 1 | Round 2 | Cause if FAIL |
|---|---|---|---|
| 1 | FAIL | FAIL (5/6) | Number/text mismatch not listed as a cause (executor slip; skill text is correct) |
| 2 | PASS | PASS (stands) | |
| 3 | PASS | PASS (stands) | |
| 4 | PASS | PASS (stands) | |
| 5 | FAIL | **PASS** (5/5) | |
| 6 | PASS | PASS (stands) | |

Possible further fix for eval 1: in SKILL.md, make the #N/A rule an output requirement, e.g. "List all of these in the answer even if the user's data is names: ... number stored as text ...". Optionally add an "#N/A checklist" line to the answer-format Watch-out section, so the executor cannot drop it.

## Round 3 (after round-2 fixes)

Fixed package: `scratchpad/blind/pkg3-sheets-formula-coach/`. New blind answer for eval 1: `scratchpad/blind/os-out-r3/sheets-formula-coach-1.md`. Grading is as strict as rounds 1 and 2.

### Eval 1 - Google Sheets VLOOKUP #N/A, decimal comma locale - PASS (6/6)

| # | Expectation | Result | Notes |
|---|---|---|---|
| 1 | Likely causes: missing FALSE, extra spaces, number/text mismatch, separator note | PASS | The "Watch out" list names all six causes: no FALSE, extra spaces, spelling (Zeynab vs Zeynəb, Latin i vs ı, now correct), value really missing, **number stored as text (or the reverse)**, and a wrong range or the name in the wrong column. The separator is covered in the first line ("uses `1,5`, so it wants `;`") and again under "Formula parse error". The round-2 gap is closed. |
| 2 | Fixed formula with `;` | PASS | `=VLOOKUP(A2;Results!A:B;2;FALSE)` and `=VLOOKUP(TRIM(A2);Results!A:B;2;FALSE)` are both valid Sheets syntax. `=XLOOKUP(A2;Results!A:A;Results!B:B;"not found")` is correct and exact by default. The comma version is offered only as a fallback ("if your Sheets accepts commas"). That is a little vague but not wrong. |
| 3 | Each piece explained | PASS | A2, `Results!A:B` (including what `!` means and that the search uses the first column), `2`, `FALSE` (with `0` as an equivalent), TRIM (typed side only, with the clean-the-list steps), IFNA and XLOOKUP. |
| 4 | Traced example table | PASS | Murad 60, Ali 70, Zaur 85, Aysel 90, with no header row. Zaur gives 85 and Leyla (not in the list) gives #N/A, both correct. "Ali " with a trailing space typed in A2 gives #N/A, also correct. The trap now sits on the typed side, so the round-2 header-row problem cannot happen. The paste-in test data matches the table. |
| 5 | Suggests IFNA only after the real cause is fixed | PASS | "When it works ... `=IFNA(VLOOKUP(A2;Results!A:B;2;FALSE);"not found")`" uses `;`. It adds "only after the plain version works, so it does not hide the real problem" and "IFNA hides only 'not found'". IFERROR is never offered. |
| 6 | Short, phone-readable | PASS (borderline) | 535 words, more than round 2 (441). It reads easily on a phone: short bold sections, a numbered cause list and one-line code blocks. The table appears twice (once as the example, once as paste-in data), and that repeat could be cut. |

Factual checks: in Sheets the menu path is Edit → Paste special → Values only. XLOOKUP is in Sheets but not in Excel 2016 or 2019, which is correct. In this locale the decimal comma goes inside formulas. "Formula parse error → check the separator first" is the correct Sheets symptom. No wrong facts.

### Skill package check: #NAME? vs wrong separator

There is no contradiction now. The three places agree:
- SKILL.md, Locale: "Wrong separator: Excel shows a 'There's a problem with this formula' dialog; Sheets shows 'Formula parse error'. `#NAME?` means a wrong, untranslated or unavailable function name (or text without quotes)."
- SKILL.md, Fixing errors, #NAME?: "typo, missing quotes around text, wrong or untranslated locale name, function not in this version. (A wrong separator gives a dialog in Excel and 'Formula parse error' in Sheets.)" The round-2 "or separator" wording is gone.
- references/errors.md: the #NAME? table has no separator row. The separator appears only under "Google Sheets: 'Formula parse error'". CHANGELOG states the same rule.

The SKILL.md #N/A rule now requires every cause "even when the data is names", including number stored as text, and IFNA after the fix. The answer follows it.

Still open, harmless: locale-and-versions.md keeps † on long-standing RU/TR names such as ЕСНД, EĞERYOKSA, #ИМЯ? and #AD?. SKILL.md tells the executor not to say these are unverified, so this does not affect answers.

### Updated total

**Pass: 6/6.** Evals 2, 3, 4 and 6 stand from round 1, eval 5 from round 2, and eval 1 passes now. **No FAIL left.**

| Eval | Round 1 | Round 2 | Round 3 | Cause if FAIL |
|---|---|---|---|---|
| 1 | FAIL | FAIL (5/6) | **PASS** (6/6) | |
| 2 | PASS | PASS (stands) | PASS (stands) | |
| 3 | PASS | PASS (stands) | PASS (stands) | |
| 4 | PASS | PASS (stands) | PASS (stands) | |
| 5 | FAIL | PASS (5/5) | PASS (stands) | |
| 6 | PASS | PASS (stands) | PASS (stands) | |
