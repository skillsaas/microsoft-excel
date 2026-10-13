# Automating Excel itself

Use these when you need Excel's own engine, or features file libraries cannot handle: real pivots, spill formulas, charts in existing files, cloud files.

| Surface | Runs where | Needs |
| :--- | :--- | :--- |
| Office.js (`Excel.run`) | Add-in in Excel for Windows, Mac, web, iPad | Add-in manifest and sideload |
| Office Scripts (TypeScript) | Excel for the web; Windows and Mac in supported Microsoft 365 builds; Power Automate | Business or education license; admin may disable it |
| Microsoft Graph `/workbook` | Files in OneDrive or SharePoint | Entra app, delegated or app token |
| COM via `pywin32` or `xlwings` | Windows with desktop Excel (xlwings also on Mac) | Excel installed, interactive session |
| VBA | Desktop Excel, Windows and Mac | Macro-enabled file, trust settings |
| Power Query (M) | Excel for Windows; Mac has a reduced connector set | Refresh in Excel |

Office Scripts do not replace every VBA capability: there is no file system access, no UI events and no other Office apps.

## Office.js
```javascript
await Excel.run(async (ctx) => {
  const ws = ctx.workbook.worksheets.getActiveWorksheet();
  const used = ws.getUsedRange();
  used.load("address,rowCount");
  await ctx.sync();                                  // data exists only after sync
  ws.getRange("F2").formulas = [["=SORT(A2:A20)"]];  // real dynamic array, spills
  await ctx.sync();
});
```
Load everything you need, then sync once. Calling `sync` inside a loop is the most common performance bug. `range.values` must match the range shape exactly.

## Office Scripts
```typescript
function main(workbook: ExcelScript.Workbook) {
  const ws = workbook.getWorksheet("Sales");
  if (!ws) throw new Error("Sheet 'Sales' not found");
  const table = ws.getTables()[0];
  if (!table) throw new Error("No table on 'Sales'");
  const vals = table.getColumnByName("Amount").getRangeBetweenHeaderAndTotal().getValues();
  const total = vals.reduce((s, [v]) => s + (typeof v === "number" ? v : 0), 0);
  ws.getRange("F2").setValue(total);
}
```

## Microsoft Graph
- Scopes: use the narrowest that works. Prefer delegated `Files.ReadWrite` on the user's own file over `Sites.ReadWrite.All`, and never log tokens.
- Session: `POST /me/drive/items/{item-id}/workbook/createSession` with body `{"persistChanges": true}`, then send the header `workbook-session-id: <id>` on each call, and close it with `POST .../workbook/closeSession`.
- Read: `GET /me/drive/items/{id}/workbook/worksheets/{name}/range(address='A1:D10')`.
- Write: `PATCH` the same URL with `{"values": [[1,2],[3,4]]}` or `{"formulas": [["=SUM(A1:A2)"]]}`.
- Tables: `GET .../workbook/tables/{name}/rows`, `POST .../workbook/tables/{name}/rows/add` with `{"values": [[...]]}`.
- Recalculate: `POST .../workbook/application/calculate` with `{"calculationType": "Full"}`.
- Errors: 429 and 503 return `Retry-After`. Wait that many seconds and retry with backoff. Large workbooks can time out; split the work into smaller ranges.

## COM (Windows) and xlwings
```python
import xlwings as xw                      # pip install xlwings; needs desktop Excel
app = xw.App(visible=False, add_book=False)
app.display_alerts = False
try:
    app.api.AutomationSecurity = 3        # msoAutomationSecurityForceDisable: block macros on open
    wb = app.books.open(r"C:\work\copy.xlsx", update_links=False)
    wb.app.calculate()
    print(wb.sheets["Sales"].range("F2").value)
    wb.save()
finally:
    app.quit()
```
Excel locks files it has open (`~$name.xlsx`). Close the workbook, or ask the user to, before editing it with another tool.

## VBA (only when the user asks for a macro)
```vba
Sub ExportSheetsToCsv()
    Dim ws As Worksheet
    On Error GoTo Done
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Copy
        ActiveWorkbook.SaveAs ThisWorkbook.Path & Application.PathSeparator & ws.Name & ".csv", xlCSVUTF8
        ActiveWorkbook.Close SaveChanges:=False
    Next ws
Done:
    Application.ScreenUpdating = True
End Sub
```
`xlCSVUTF8` (62) needs Excel 2019 or 365. In Excel 2016, use `xlCSV` (6) and handle the encoding yourself.
Without Excel, you can write VBA source for the user to paste, but you cannot add or run it in a file. openpyxl only preserves an existing `vbaProject.bin`.

## Power Query (M)
```m
let
    Source = Csv.Document(File.Contents("C:\data\sales.csv"), [Delimiter=",", Encoding=65001]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Promoted,
        {{"Date", type date}, {"Account ID", type text}, {"Amount", type number}}, "en-US")
in
    Typed
```
Always pass a culture to type conversions. Queries are stored in `customXml` and the data model. File libraries cannot run or edit them.
