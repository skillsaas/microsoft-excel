# Excel formula reference for agents

## Choose the function
Availability: XLOOKUP, FILTER, SORT, UNIQUE, SEQUENCE, LET in Excel 2021+ and 365. LAMBDA helpers, TEXTSPLIT, VSTACK, TAKE, DROP in 365 and 2024. GROUPBY, PIVOTBY (365 Current Channel since Sep 2024; Microsoft's page also lists 2021 and 2024, treat as unverified) and REGEX* (365 only). If the reader's version is unknown, use the right column.

| Goal | Modern | Compatible with 2016/2019 |
| :--- | :--- | :--- |
| Lookup exact | `XLOOKUP(k, keys, vals, "none")` | `INDEX(vals, MATCH(k, keys, 0))` |
| Lookup last match | `XLOOKUP(k, keys, vals, , 0, -1)` | `LOOKUP(2, 1/(keys=k), vals)` |
| Two-key lookup | `XLOOKUP(1, (a=k1)*(b=k2), vals)` | `INDEX(vals, MATCH(1, (a=k1)*(b=k2), 0))` entered with Ctrl+Shift+Enter |
| Conditional sum | `SUMIFS(vals, crit_rng, crit)` | same |
| Distinct list | `UNIQUE(rng)` | pivot table or Remove Duplicates |
| Filter rows | `FILTER(tbl, cond, "")` | AutoFilter / Advanced Filter |
| Group and sum | `GROUPBY(keys, vals, SUM)` | `SUMIFS` over `UNIQUE`, or pivot |
| Split text | `TEXTSPLIT(t, ",")` | Text to Columns, `MID`/`FIND` |
| Join text | `TEXTJOIN(", ", TRUE, rng)` | `TEXTJOIN` (2019) or `CONCATENATE` (2016) |
| Running total | `SCAN(0, rng, LAMBDA(a,x,a+x))` | `=SUM($B$2:B2)` filled down |
| Regex (365 only) | `REGEXEXTRACT`, `REGEXTEST`, `REGEXREPLACE` | none; use Power Query or VBA |

## Error values and what to do
| Error | Usual cause | Agent fix |
| :--- | :--- | :--- |
| `#N/A` | Lookup key not found; trailing spaces; number vs text | `TRIM`/`VALUE` keys, or `XLOOKUP(..., "not found")` |
| `#VALUE!` | Text in arithmetic, wrong argument type | Convert text numbers; check ranges are the same size |
| `#REF!` | Deleted rows or columns, broken external link | Restore from the original; re-point the reference |
| `#NAME?` | Typo, missing defined name, function unsupported in this Excel or LibreOffice | Use the compatible column above |
| `#DIV/0!` | Divide by zero or blank | `IFERROR(x, 0)` only if the user agrees zero is right |
| `#SPILL!` | Spill range not empty, or the formula is in a Table | Clear the range or move the formula out of the Table |
| `#CALC!` | Empty array (FILTER with no matches) | Add the `if_empty` argument |
| `#NUM!` | Invalid numeric input or result out of range | Check inputs and iteration settings |
| `#NULL!` | Space used as intersection of non-intersecting ranges | Use `,` or `:` |

## Input and expected result
| Write into the cell | Data | Shows |
| :--- | :--- | :--- |
| `=XLOOKUP("B",A2:A4,B2:B4,"none")` | A2:A4 = A,B,C; B2:B4 = 1,2,3 | 2 |
| `=SUMIFS(C:C,A:A,"EMEA",B:B,">=2026-01-01")` | Dates as real dates | sum of matching rows |
| `=TEXT(A2,"yyyy-mm-dd")` | A2 = 46296 | 2026-10-01 |
| `=ROUND(0.1+0.2,10)=0.3` | | TRUE |

## Locale
Formula separators follow the locale in the UI (`;` in de-DE, fr-FR) but files always store the en-US form with `,`. When writing formulas through openpyxl, Office.js `range.formulas`, or Graph, always use English function names and `,`. Office.js `formulasLocal` is the exception.

## Numbers and dates
- Prefer library conversion (`openpyxl.utils.datetime.from_excel(serial, wb.epoch)`). By hand, in the 1900 system: `(d - date(1899, 12, 30)).days`, valid from 1900-03-01. In the 1904 system the serial is 1,462 smaller.
- `TEXT(A1, "yyyy-mm-dd")` output depends on locale codes in some languages; prefer number formats over `TEXT` for display.
