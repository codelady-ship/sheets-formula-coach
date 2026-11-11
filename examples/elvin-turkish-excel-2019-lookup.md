# Example: lookup in Turkish Excel 2019

Fictional user and data. Shows version check, translated names and a traced example.

## Request

> My Excel 2019 is in Turkish. Sheet "Fiyat" has product codes in A and prices in C. In my order sheet the code is in B2 and I want the price in D2. A friend sent me =XLOOKUP(B2,Fiyat!A:A,Fiyat!C:C) but I get #AD?.

## Result

**Why:** XLOOKUP (Turkish: `ÇAPRAZARA`) is **not in Excel 2019**, so Excel shows `#AD?` (= #NAME?). Also, Turkish Excel uses `;` and Turkish names.

```
=İNDİS(Fiyat!$C$2:$C$500;KAÇINCI(KIRP(B2);Fiyat!$A$2:$A$500;0))
```

- **Where:** D2, then drag down.
- **How it works:**
  - `KIRP(B2)` (TRIM) removes stray spaces from the code.
  - `KAÇINCI(…;0)` (MATCH) finds the row of the code; `0` = exact match.
  - `İNDİS(…)` (INDEX) returns the price from that row of column C.
  - After it works, you can show "bulunamadı" for missing codes: `=EĞERYOKSA(İNDİS(…);"bulunamadı")` (IFNA, hides only "not found").
  - `$` keeps the ranges fixed when you drag.

**Example (sheet Fiyat)**
| A Code | C Price |
|---|---|
| P-101 | 12,50 |
| P-102 | 8 |
| P-205 | 30 |

B2 = `P-102 ` (with a trailing space) → KIRP → `P-102` → KAÇINCI = 2 → İNDİS → **8**.

**Watch out**
- Codes stored as numbers in one sheet and as text in the other still give #YOK (#N/A): make both the same type.
- Names look different? Type `KAÇ` and pick from the autocomplete list.

**English / comma version** (Excel 365 or Sheets in a US-style locale):
```
=XLOOKUP(TRIM(B2),Fiyat!A:A,Fiyat!C:C,"not found")
```

Explain more, or test data to paste?
