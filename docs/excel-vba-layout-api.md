# Excel VBA Layout Query Surface

This is the checked-in inventory for Excel layout-query APIs we use as parity oracles. Mirrors `docs/word-vba-layout-api.md`.

## Microsoft Excel APIs

- `Range.Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)` returns the A1 or R1C1 reference of a `Range`. Useful for round-tripping cell positions across automation/browser/native boundaries.
- `Range.Row`, `Range.Column`, `Range.Rows.Count`, `Range.Columns.Count` expose 1-based row/column indices and cardinalities for a range.
- `Range.Top`, `Range.Left`, `Range.Width`, `Range.Height` return geometry in points relative to the worksheet origin. These are the Excel equivalents of Word's `Range.Information(wdHorizontalPositionRelativeToPage)`/`wdVerticalPositionRelativeToPage`.
- `Range.Value` and `Range.Value2` return cell values; `Range.Formula`, `Range.FormulaR1C1`, and `Range.HasFormula` expose the underlying formula text.
- `Range.NumberFormat`, `Range.NumberFormatLocal`, `Range.Style` expose number format strings and style names.
- `Range.MergeCells` is `True` when the range is merged. `Range.MergeArea` returns the merged range, useful for finding the primary cell.
- `Range.WrapText`, `Range.HorizontalAlignment`, `Range.VerticalAlignment`, `Range.Orientation` expose layout flags.
- `Range.EntireRow.Hidden`/`EntireColumn.Hidden` and `Range.RowHeight`/`Range.ColumnWidth` expose row/column sizing (points for row height, character units for column width — converted with `Worksheet.StandardWidth`).
- `Window.RangeFromPoint(x, y)` returns the range under a screen-pixel position. Like Word's `Window.GetPoint`, this is fragile in the headless/scheduled-task path because Excel raises an error when the worksheet view isn't fully visible.
- `Application.CentimetersToPoints`/`InchesToPoints` and the inverse helpers convert units; our comparator converts Excel points to CSS pixels with `96 / 72`.
- `Worksheet.UsedRange` and `Worksheet.Names` expose worksheet-wide bounds and defined names. `Workbook.Sheets`/`Sheets(index)` enumerate sheets.

Primary docs:

- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.address>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.top>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.left>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.width>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.height>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.value>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergecells>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat>
- <https://learn.microsoft.com/en-us/office/vba/api/excel.window.rangefrompoint>

## Implemented Here

- `tools/excel-parity/extract-excel-layout.vbs` emits the stable Excel-range fields above for every non-empty cell in every visible sheet of an XLSX, plus merge geometry.
- Planned: `RedBat.Spreadsheet.API.getLayoutCells` to return per-cell geometry alongside text/format/value.
- Planned: `RedBat.Spreadsheet.API.getRangeInformation(editor, sheetId, sref)` mirroring Excel's `Range`-level layout fields for a single range.
- Planned: `RedBat.Spreadsheet.API.computeRangeStatistics(editor, sheetId, sref)` returning cell/row/column/word/character counts for a single range.
- Planned: `RedBat.Spreadsheet.API.getRangeBounds(editor, sheetId, sref)` returning sheet-relative geometry analogous to Word's `getRangeBounds`.
- Planned: `tools/excel-parity/compare-vba-layout-results.mjs` to align Excel TSV output and browser results into per-document layout diffs.

## Not Yet Modeled

- Screen-pixel `Window.RangeFromPoint` behavior, including visibility failures and zoom effects.
- Conditional format/style resolution at the cell level (Excel exposes `Range.DisplayFormat`; we currently snapshot the raw `Range.NumberFormat` only).
- Chart, pivot, and shape geometry (`Chart.Top`, `PivotTable.RowGrand`, `Shape.LockAspectRatio`).
- Multi-sheet `Range.Information`-style probes. Current API support is single sheet plus an explicit range reference.
- Workbook-level protection and shared-workbook state.
