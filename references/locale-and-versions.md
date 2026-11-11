# Locale, translated names and versions

Checked on Microsoft and Google help pages on **2026-11-11**. Microsoft's "Functions Translator" add-in is **discontinued** ("no longer available and no longer supported"), so don't send users to it.

## Separators

- **Excel** uses the **list separator from the Windows regional settings** (Control Panel → Region → Additional settings), unless "Use system separators" is turned off in File → Options → Advanced. US English uses `,`. Where the decimal sign is a comma (`1,5`), the list separator is usually `;` (Azerbaijan, Türkiye, Russia, most of Europe).
- **Google Sheets** follows File → Settings → Locale. If the locale writes `1,5`, use `;`.
- Quick test for the user: "Does one and a half look like `1.5` or `1,5` in your sheet?" `1,5` → use `;`.
- In `;` locales, decimals inside formulas use a comma: `=A2*1,5`.
- In Sheets, array literals change too: `{1,2}` becomes `{1\2}` in many comma-decimal locales. If an array literal fails, try `\` between columns.

## Translated function names (Excel)

Excel shows function names in the interface language. Russian and Turkish Excel use translated names; **Azerbaijani Excel** is only a partial language pack and there is no evidence that it translates function names, so Azerbaijani users usually see **English, Russian or Turkish** names: ask which one their Excel shows.

Tip that always works: type the first letters and **pick the name from the autocomplete list**.

| English | Russian | Turkish |
|---|---|---|
| VLOOKUP | ВПР | DÜŞEYARA |
| XLOOKUP | ПРОСМОТРX | ÇAPRAZARA |
| INDEX | ИНДЕКС | İNDİS |
| MATCH | ПОИСКПОЗ † | KAÇINCI |
| IF | ЕСЛИ | EĞER |
| IFERROR | ЕСЛИОШИБКА | EĞERHATA |
| IFNA | ЕСНД † | EĞERYOKSA † |
| SUMIF | СУММЕСЛИ | ETOPLA |
| SUMIFS | СУММЕСЛИМН | ÇOKETOPLA |
| COUNTIF | СЧЁТЕСЛИ † | EĞERSAY † |
| COUNTIFS | СЧЁТЕСЛИМН | ÇOKEĞERSAY |
| AVERAGEIFS | СРЗНАЧЕСЛИМН † | ÇOKEĞERORTALAMA † |
| SUM | СУММ † | TOPLA † |
| AND / OR | И / ИЛИ † | VE / YADA † |
| FILTER | ФИЛЬТР | FİLTRE |
| UNIQUE | УНИК † | BENZERSİZ † |
| SORT | СОРТ † | SIRALA † |
| DATE | ДАТА | TARİH |
| TODAY | СЕГОДНЯ † | BUGÜN † |
| EOMONTH | КОНМЕСЯЦА | SERİAY |
| TRIM | СЖПРОБЕЛЫ | KIRP |
| TEXTBEFORE | ТЕКСТДО | ÖNCEKİMETİN |
| VALUE | ЗНАЧЕН † | SAYIYAÇEVİR † |
| PMT | ПЛТ † | DEVRESEL_ÖDEME † |
| TRUE / FALSE | ИСТИНА / ЛОЖЬ | DOĞRU / YANLIŞ |
| #N/A | #Н/Д | #YOK |
| #NAME? | #ИМЯ? † | #AD? † |

Sources: Microsoft Support alphabetical function lists in ru-ru and tr-tr, and single-function pages (VLOOKUP, XLOOKUP, COUNTIFS, EOMONTH, TRIM, MATCH). † = widely used name, not re-checked on Microsoft's page on 2026-11-11: tell the user to confirm it with autocomplete. In ПРОСМОТРX the last letter may be Latin X or Cyrillic Х depending on the source; picking from autocomplete avoids the problem.

Format strings in TEXT() are also local (e.g. Russian `"ДД.ММ.ГГГГ"` vs English `"dd.mm.yyyy"`). If TEXT() shows the format code instead of a date, the codes are in the wrong language.

## Google Sheets names

Sheets can show function names in English or in the display language: File → Settings → **"Always use English function names"** (official, Google Docs Editors Help). If the user's Sheets shows local names, tick that box to use English names and the formulas in this skill.

## Version check (Excel)

| Function / feature | Works in |
|---|---|
| XLOOKUP, XMATCH | 365, 2021, 2024, mobile. **Not in 2016 or 2019** (official) → use INDEX + MATCH |
| FILTER, UNIQUE, SORT, SEQUENCE (dynamic arrays) | 365, 2021, 2024 → older: pivot table, Advanced Filter, or helper columns |
| LET | 365, 2021, 2024 |
| LAMBDA | 365, 2024 |
| TEXTBEFORE, TEXTAFTER, TEXTSPLIT, VSTACK | 365, 2024 → older: LEFT/MID/FIND |
| GROUPBY, PIVOTBY | 365 only → older: pivot table or SUMIFS grid |
| REGEXEXTRACT, REGEXTEST, REGEXREPLACE | 365 only (Sheets has its own REGEX functions) |
| Checkboxes (Insert → Checkbox) | 365 and Excel for the web (snippet) → older: Yes/No drop-down |
| IFS, MAXIFS, MINIFS, TEXTJOIN | 2019 and newer |

Google Sheets has XLOOKUP, FILTER, UNIQUE, SORT, LET, LAMBDA, QUERY, REGEX functions and checkboxes.

How the user finds their version: Excel → File → Account → About Excel. "Microsoft 365" = newest functions.
