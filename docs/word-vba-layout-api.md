# Word VBA Layout Query Surface

This is the checked-in inventory for layout-query APIs we use as parity oracles.

## Microsoft Word APIs

- `Range.Information(Type)` and `Selection.Information(Type)` return `WdInformation` values for a range or selection. The layout-relevant constants include page number, adjusted page number, section number, page count, horizontal/vertical position relative to page, horizontal/vertical position relative to text boundary, first character line/column number, table row/column numbers, maximum table rows/columns, end-of-row marker, content control, field code/result, and table membership.
- `Range.ComputeStatistics(Statistic)` returns counts for a range. The useful `WdStatistic` values here are lines, pages, paragraphs, words, characters, and characters-with-spaces. Word notes that table cell ranges need the end-of-cell marker removed for accurate line counts.
- `Window.GetPoint(left, top, width, height, obj)` returns screen-pixel bounds for a visible `Range` or `Shape`. It is useful for interactive probes, but fragile in the current headless/scheduled-task path because Word raises an error when the object is not fully visible.
- `Cell.RowIndex`, `Cell.ColumnIndex`, `Cell.Width`, and `Cell.Height` expose table-local geometry and indices. We use row/column indices now; width/height are still better measured from rendered/painted boxes on our side because Word's auto-height cells can return `wdUndefined`.
- `Application.PointsToPixels`, `PixelsToPoints`, and related unit conversion methods are useful when comparing `Information` point coordinates with browser/native pixel coordinates. Our comparison scripts convert Word points to CSS pixels with `96 / 72`.

Primary docs:

- <https://learn.microsoft.com/en-us/office/vba/api/word.range.information>
- <https://learn.microsoft.com/en-us/office/vba/api/word.selection.information>
- <https://learn.microsoft.com/en-us/office/vba/api/word.wdinformation>
- <https://learn.microsoft.com/en-us/office/vba/api/word.range.computestatistics>
- <https://learn.microsoft.com/en-gb/office/vba/api/word.wdstatistic>
- <https://learn.microsoft.com/es-es/office/vba/api/word.window.getpoint>
- <https://learn.microsoft.com/en-us/office/vba/api/word.cell.columnindex>
- <https://learn.microsoft.com/en-us/office/vba/api/word.cell.width>

## Implemented Here

- `tools/word-parity/extract-word-layout.vbs` now emits the stable `WdInformation` values listed above plus `ComputeStatistics(wdStatisticLines)` and `ComputeStatistics(wdStatisticPages)` for each emitted line range.
- `Rbt.Word.API.getLayoutLines` now includes VBA-style fields alongside existing line geometry: page, adjusted page, section, page count, page-relative x/y, text-boundary-relative x/y, first-character line/column, table row/column/max row/max column, content-control and field-code/result flags.
- `Rbt.Word.API.getVbaLayoutLines` is an explicit alias for the enriched line query.
- `Rbt.Word.API.getRangeInformation(editor, blockId, startOffset, endOffset)` mirrors the useful `Range.Information` subset for a paragraph range.
- `Rbt.Word.API.getSelectionInformation(editor)` mirrors `Selection.Information` for the active end of the current selection.
- `Rbt.Word.API.computeRangeStatistics(editor, blockId, startOffset, endOffset)` returns line/page/paragraph/word/character counts for a single paragraph range.
- `Rbt.Word.API.getRangeBounds(editor, blockId, startOffset, endOffset)` returns document/page bounds for the range, analogous to the stable document-coordinate part of `Window.GetPoint`.
- `tools/word-parity/compare-word-browser-layouts.mjs` carries the added Word/browser fields through `layout-diff.tsv`.

## Not Yet Modeled

- Screen-pixel `Window.GetPoint` behavior, including visibility failures and scroll/zoom effects.
- Header/footer, footnote/endnote, citation, bibliography, clipboard, keyboard, and mail-pane `WdInformation` flags.
- Multi-block `Range.Information` and `ComputeStatistics` semantics. Current API support is single paragraph range plus current selection active-end information.
- Exact Word section numbering across every section-break variant. Our current section value follows top-level paragraph section breaks and is sufficient for finding gross section/page-origin mistakes.
