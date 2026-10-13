# Recipes (openpyxl 3.1, xlsxwriter 3.2, pandas 2.2+ or 3.x)

Blocks were tested with openpyxl 3.1.5, xlsxwriter 3.2.9, pandas 3.0, msoffcrypto-tool 6.0 and python-calamine (exact versions and dates: CHANGELOG.md). Work on a copy of the user's file.

## What survives an openpyxl load and save
Tested with openpyxl 3.1.5 (plus Pillow) on a workbook written by xlsxwriter 3.2.9, not on a file saved by Excel. Excel-authored files (chart styles, `extLst` extensions, pivot caches, slicers) and other openpyxl versions can behave differently, so inventory before and after every edit (`inspect_xlsx.py --json`, compare the reports).

| Feature | Result in the test | Treat as |
| :--- | :--- | :--- |
| Values, formulas, styles, number formats, freeze panes, autofilter, hidden rows | kept | usually safe |
| Tables, defined names, data validation, conditional formats (cell rules, data bars) | kept | usually safe; Excel 2010+ extension rules (`extLst`) were not tested |
| Comments | kept, rewritten as new parts | text kept; size and position may change |
| Images | kept only with Pillow installed | may be lost |
| Charts | kept but rebuilt by openpyxl (chart XML changed) | kept but altered: formatting may change; Excel-authored charts may lose more |
| Cached formula values | **dropped** (`data_only` reads `None` after save) | recalculate before reading values |
| Pivot tables and caches, slicers, timelines, sparklines, form controls, ActiveX, unknown parts | not tested | may be lost or altered; edit through Excel instead |

If the workbook has charts, images, pivots, slicers or unusual parts and the user cannot accept a change to them, do not save it with openpyxl. Use Excel automation, or keep the original and report the limit.

## Pivot tables
openpyxl **cannot create** pivot tables. It has limited read and preservation support for existing ones (it can load and write back the pivot and cache objects), but no reliable general editing API, and pivots were not in our round-trip test. Treat edits to pivot definitions, caches, slicers and timelines as unsupported and use Excel automation. Options:
1. Need a summary and not a live pivot: compute it with pandas and write the values (below), or use `SUMIFS` formulas, or `GROUPBY`/`PIVOTBY` if every reader has Microsoft 365. LibreOffice lacks GROUPBY and PIVOTBY, so `recalc.py` turns them into `#NAME?`; do not use them when values must come from recalc.
2. Need a real PivotTable: create it through Excel, using Office.js `worksheet.pivotTables.add(name, sourceRange, destRange)`, COM/xlwings, or by hand.
```python
import pandas as pd
df = pd.read_excel("sales.xlsx", sheet_name="Sales")
pv = df.pivot_table(index="Region", columns=df["Date"].dt.year, values="Amount", aggfunc="sum", fill_value=0)
with pd.ExcelWriter("sales.edited.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
    pv.to_excel(xw, sheet_name="Summary")
```
`mode="a"` keeps the other sheets, but openpyxl still rewrites the whole file, so the preservation limits in SKILL.md apply.

## Charts (openpyxl)
```python
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
wb = load_workbook("sales.edited.xlsx")
ws = wb["Summary"]
chart = BarChart()
chart.title, chart.y_axis.title = "Amount by region", "Amount"
data = Reference(ws, min_col=2, min_row=1, max_col=ws.max_column, max_row=ws.max_row)
cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "H2")
wb.save("sales.edited.xlsx")
```
Charts already in a file that openpyxl loads are kept but rebuilt on save (see the table above): in our xlsxwriter-authored test the chart survived with rewritten XML. Excel-authored charts with styles, combo types or extensions may lose formatting or parts. If the file has charts (`inspect_xlsx.py` shows `charts > 0`) and their look matters, edit with Excel or write the changes through Graph instead.

## Data validation (dropdown)
```python
from openpyxl.worksheet.datavalidation import DataValidation
dv = DataValidation(type="list", formula1='"EMEA,APAC,AMER"', allow_blank=True,
                    showErrorMessage=True, error="Pick a region")
ws.add_data_validation(dv)
dv.add("B2:B1000")
```
An inline list must stay under 255 characters. For longer lists, point to a range: `formula1="=Lists!$A$2:$A$50"`.

## Conditional formatting
```python
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import PatternFill
ws.conditional_formatting.add("D2:D1000",
    CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.conditional_formatting.add("E2:E1000", ColorScaleRule(start_type="min", start_color="F8696B",
                                                         end_type="max", end_color="63BE7B"))
```

## Named ranges
```python
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names["TaxRate"] = DefinedName("TaxRate", attr_text="Settings!$B$1")   # openpyxl 3.1 API
ws["G2"] = "=D2*TaxRate"
```
In openpyxl before 3.1, use `wb.defined_names.append(DefinedName(...))`. Sheet-scoped names live in `ws.defined_names`.

## Merged cells
```python
for rng in list(ws.merged_cells.ranges):
    top_left = ws.cell(rng.min_row, rng.min_col).value
    ws.unmerge_cells(str(rng))          # only if the user wants a flat table
    for row in ws.iter_rows(min_row=rng.min_row, max_row=rng.max_row,
                            min_col=rng.min_col, max_col=rng.max_col):
        for c in row:
            c.value = top_left
```
When reading, fill merged values forward as above (on a copy) or skip the non-top-left cells. Writing into a non-top-left merged cell raises an error in openpyxl.

## Sheet and workbook protection
Sheet protection is not encryption: openpyxl reads and edits protected sheets.
```python
ws.protection.sheet = True
ws.protection.password = "user-supplied"      # stored as a hash; weak, deters edits only
from openpyxl.styles import Protection
ws["B2"].protection = Protection(locked=False)   # editable input cell
wb.security.lockStructure = True              # workbook structure (add or delete sheets)
```
Do not remove protection unless the user owns the file and asks.

## Encrypted (password to open) files
```python
import io, msoffcrypto                      # pip install msoffcrypto-tool
from openpyxl import load_workbook
with open("locked.xlsx", "rb") as f:
    office = msoffcrypto.OfficeFile(f)
    office.load_key(password=user_password)   # only a password the user gave you
    buf = io.BytesIO()
    office.decrypt(buf)
wb = load_workbook(buf)
```
Do not write the decrypted copy to shared locations. To re-encrypt, use `msoffcrypto.format.ooxml.OOXMLFile(f).encrypt(password, out)`, or ask the user to re-protect it in Excel.

## pandas without data loss
```python
sheets = pd.read_excel(p, sheet_name=None, dtype=str, keep_default_na=False)   # dict: sheet name -> DataFrame
df = sheets["Sales"]                                     # one sheet; or pd.concat(sheets.values()) if they share columns
df["Amount"] = pd.to_numeric(df["Amount"].str.strip().str.replace(",", ""), errors="coerce")
df = df.astype(object).where(df.notna(), None)                                 # NaN -> empty cell
```
`df.to_excel` writes values only, with no formulas or styles. Use it for new files, never to update a formatted workbook.

## New workbook with xlsxwriter
```python
import xlsxwriter
wb = xlsxwriter.Workbook("report.xlsx", {"constant_memory": True, "strings_to_formulas": False})
ws = wb.add_worksheet("Data")
ws.write_row(0, 0, ["ID", "Amount"])
ws.write_string(1, 0, "000123")
ws.write_number(1, 1, 99.5)
ws.write_formula(2, 1, "=SUM(B2:B2)", None, 99.5)   # optional cached value
wb.close()
```
- xlsxwriter cannot read or edit existing files.
- `strings_to_formulas: False` stops `write()` from treating "=..." text as a formula, which helps against injection.
- `constant_memory` requires writing rows in order.
- `write_dynamic_array_formula` writes real spill metadata, which openpyxl cannot.

## Fast reading of large files
`pip install python-calamine`, then `pd.read_excel(p, engine="calamine", dtype=str)` (pandas 2.2+). It reads xlsx, xlsb, xls and ods, but only values (no formulas or styles). Without `dtype=str`, text IDs such as "000123" came back as `123.0` in testing.
