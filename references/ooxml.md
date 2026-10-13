# .xlsx internals (ECMA-376 SpreadsheetML)

An .xlsx is a zip. Inspect with `unzip -l file.xlsx`.

```
[Content_Types].xml          # MIME type per part; missing entry = "Excel found unreadable content"
_rels/.rels                  # points to xl/workbook.xml
xl/workbook.xml              # sheets (name, sheetId, r:id), definedNames, workbookPr (date1904), calcPr
xl/_rels/workbook.xml.rels   # r:id -> worksheets/sheetN.xml
xl/worksheets/sheet1.xml     # <sheetData><row r="1"><c r="A1" t="s" s="3"><v>0</v></c>
xl/sharedStrings.xml         # <si><t>text</t></si>; t="s" cells hold the index
xl/styles.xml                # numFmts, fonts, fills, borders, cellXfs; c/@s is an index into cellXfs
xl/tables/table1.xml         # ListObject definition (ref, columns, autoFilter)
xl/calcChain.xml             # calc order; safe to delete (and its rel + content type) when editing formulas by hand
xl/externalLinks/            # links to other workbooks
xl/vbaProject.bin            # macros (only .xlsm); content type must be macroEnabled
xl/pivotTables/, xl/pivotCache/, xl/charts/, xl/drawings/
docProps/core.xml, app.xml   # author, timestamps, app version
```

Cell `t` attribute: `s` shared string, `str` formula string result, `inlineStr`, `b` boolean, `e` error, `n` or absent number. Formula: `<c r="B2"><f>SUM(A1:A9)</f><v>45</v></c>`; `<v>` is the cached value. Dynamic arrays use `<f t="array" ref="A1:A5" aca="false">` plus metadata in `xl/metadata.xml` (`cm` attribute on the cell).

Dates are numbers with a date `numFmtId` (built-in 14 to 17 and 22 are dates or date-times, 18 to 21 are times; or a custom format containing d/m/y). There is no date type in the cell itself. Built-in formats 14 to 22 display according to the reader's locale (14 shows m/d/yyyy in en-US and dd/mm/yyyy in en-GB). Use an explicit custom format such as `yyyy-mm-dd` when the layout matters.

Repair recipe for "We found a problem with some content":
1. Unzip, check every part listed in `[Content_Types].xml` exists and every rel target exists.
2. Validate XML well-formedness (`xmllint --noout`).
3. Remove `xl/calcChain.xml` and its references only when the calc chain is corrupt or inconsistent; do not remove it as routine preprocessing.
4. Check shared string indexes do not exceed `<sst count>`; style indexes do not exceed `cellXfs` count.
5. Re-zip with `[Content_Types].xml` included at the root (order does not matter, but paths are case-sensitive).
