# Template screening: rendering differences (2026-11-14)

Found while choosing documents for a demo video. About 110 Microsoft Office
templates (downloaded from `binaries.templates.cdn.office.net/support/templates/en-us/<id>_win32.dotx|xltx`,
converted to .docx/.xlsx by switching the main part's content type) and a few
government templates were opened in the browser editors (redbat's mega
frontend, `@rbt/office` 0.1.33). They were compared with Word and Excel 2019 on
`winvm` (export to PDF, page 1; Excel A1:Z60 via CopyPicture).

Evidence lives outside the repo, because Microsoft's templates aren't ours to commit:

- Files: `~/Dev/office-hype/screening/word/<id>.docx`, `~/Dev/office-hype/screening/excel/<id>.xlsx`
  (866 Word and 513 Excel templates; only about 110 were compared).
- Side-by-sides (ours left, Office right): `~/Dev/office-hype/screening/side-by-side/w-<id>.png`, `x-<id>.png`.
- Scripts used: `~/Dev/office-hype/screening/scripts` (ours.mjs renders ours, vmrun.sh
  batches Word/Excel on winvm) and `~/Dev/office-hype/screening/vm`.

Caution: the Word side of the screening side-by-sides was exported with
`word.ScreenUpdating = False`, which makes Word's PDF export drop paragraph borders
(`pBdr`) that Word did read. A border that only shows on our side there is not a
difference. `tools/word-parity/print-word-pdf.vbs` had the same line until 2026-11-14.

Status column: open / fixed (commit) / not a bug (why). Some examples were matched to
the report after the fact: if one doesn't show the difference, search the side-by-sides.

## Spreadsheet

| # | Difference | Examples | Status |
|---|---|---|---|
| S1 | Cell text is drawn one CSS px per point (11pt → 11px; Excel draws 14.67px), so all sheet text is about 25% small. Column widths and row heights match. | every workbook | fixed (0ae66a9d): 11 pt is drawn 14.67 px; rows auto-fit to Excel's line heights (20/40/60 px for 1/2/3 lines of Calibri 11) |
| S2 | A cell's font family is ignored when the face isn't installed: Georgia, Constantia, Segoe UI, Calibri, Century Gothic and Franklin Gothic all fall back to a generic sans. The Word side ships web fonts for some of these. | x-tf03986924, x-tf04035471 | fixed in the grid (7b4fa526): `SetFontLoader(family, bold, italic)` fetches the faces a workbook shows when the system lacks them, and Calibri without a face is drawn in Arial at Calibri's size (`font-size-adjust`). Needs the host to pass a loader: redbat's word-text-stack has Calibri, Georgia, Segoe UI, Cambria ...; Constantia, Century Gothic and Franklin Gothic it does not ship |
| S3 | Sparklines aren't drawn. | x-tf02802361 | open: sparklines are not read into the model (only kept as XML). A fix needs the x14:sparklineGroups parsed (tools/excel-native-runner's ChartRender.parseSparklines has a parser) and drawn per cell in the grid, canvas and PDF painters |
| S4 | Pattern (hatched) fills aren't drawn. | x-tf03107663 | open: XlsxReader.parseFill blends a pattern to one colour (`patternDensity`). A fix needs the pattern type and both colours kept in CellFormat and drawn as a tile (CSS gradient / canvas pattern / PDF) |
| S5 | Rich-text runs in a cell lose their colours ("OPEN vs. CLOSED", "Step one:"). | x-tf03784001 | fixed (7b4fa526): the web grid did not draw `RichRuns` at all, and runs in a theme colour lost it (the reader kept only `rgb`) |
| S6 | Accounting format doesn't pin `$` to the cell's left edge. | x-tf03458075 | fixed (7b4fa526): the fill spaces ran together in the DOM; the sign and the number are now laid at the cell's two edges (example: x-tf03986991, Gross profit) |
| S7 | Icon sets look different from Excel's. | x-tf02897386 | open: icon sets are drawn as Unicode characters (`ConditionalFormatting.iconGlyph`). A fix needs Excel's icons as small SVGs per set and index, used by the grid, canvas and PDF |
| S8 | `###` where Excel shows the number (attendance sheet). | x-tf02780235 | fixed (7b4fa526): a one-letter date format ("d") was .NET's short date, so "1/1/2022" didn't fit; "m"/"mm" alone are now the month too |
| S9 | A pictogram shows its alt text ("decorative element") instead of the picture. | x-tf02930056 | open: the example draws its picture; what differs is in x-tf00414392, whose `leftBrace` shapes were filled rectangles - braces are now drawn (7b4fa526). No alt-text case found among the side-by-sides |
| S10 | Chart series labels spill outside the chart; pie data labels are missing. | x-tf03783862, x-tf03986584 | fixed (7b4fa526): the "labels" were the chart's source cells under the `;;;` format, whose empty text section Excel hides (a text's format section was ignored). Pie labels show in x-tf03986584 already |
| S11 | Filter buttons show on table headers where Excel hides them (table with `showAutoFilter`/header buttons off). | x-tf03427338 | fixed (7b4fa526): `hiddenButton` on a table's filter columns is read; also x-tf03458075, x-tf03986584, the attendance calendar |
| S12 | A General number too wide for its cell shows `4E+...`; Excel shrinks it to `4E+05`. | `=SUM` of salaries in a narrow column | fixed (0ae66a9d): General and #### use the column's whole-pixel width; a number is never ellipsised |
| S13 | The browser importer drops `ht` on rows without `customHeight`. Excel saves auto-fitted rows that way, so wrapped rows from Excel-saved files open at the default 20px. Fixing it changes what export writes. | any Excel-saved workbook with wrapped text | open: the reader drops a non-custom `ht` (XlsxReader.rowFromNode). Using it makes export write those rows as customHeight="1", so a fix needs the model to know which heights are auto-fitted |
| S14 | Chart frame: Excel draws the chart area's border with rounded corners (`c:roundedCorners` absent or 1, the default spPr line); we draw no frame. | ~/Dev/office-hype/docs/designed/dashboard.xlsx (both charts) | fixed (7b4fa526): a chart part without `spPr` has Excel's grey frame; `roundedCorners` absent means rounded |
| S15 | A line chart's value axis loses its default major gridlines (Excel draws them); the category axis skips every other label where Excel fits all twelve months. | dashboard.xlsx ("Fleet total") | fixed (7b4fa526): grid lines and axis lines without a line of their own are black (Excel 2019); category labels are thinned only when they would touch; `smooth` lines are curves |
| S16 | Chart titles draw larger and regular; Excel's default title here is bold and smaller (`c:title` without txPr: 14pt bold in Excel 2019's default chart style for openpyxl's `style` 10/12). | dashboard.xlsx | fixed (7b4fa526): a title of its own text that names no size is 10 pt bold, black; a title or legend without `overlay` is laid over the plot, as Excel 2019 does |
| S17 | Data bars without the x14 extension draw solid; Excel draws them with the 2007 gradient (colour fading to white). | dashboard.xlsx (Total column) | fixed (7b4fa526): a data bar rule without the x14 part fades its colour to white |

## Word

| # | Difference | Examples | Status |
|---|---|---|---|
| W1 | Shapes in groups or anchored to the page often go missing: coloured panels, photos inside groups, logo blocks, background art. | w-tf16412190, w-tf16392850, w-tf16392545, w-tf56355079, w-tf04011831, w-tf77989015, w-tf56090272 | fixed (cd98a6f6): DrawingML colours were read from a:srgbClr only, so a:schemeClr fills (most template art) became white; theme colours with lumMod/lumOff/tint/shade, wps:style fillRef/lnRef and a:blipFill pictures are now read. Still open in tf04011831 (certificate): the corner ornaments (a:custGeom paths with a:arcTo, which DocxReader.parseCustom drops) and the group’s hatched border picture are not drawn, nor the filled name pills; placement differences are under W3 |
| W2 | A group's fill or outline loses its colour. | w-tf02933929 (candles), w-tf22897213, w-tf00241780 | fixed (cd98a6f6): same cause as W1 (theme colours of group shapes) |
| W3 | Text boxes are placed wrongly; one brochure's text overlaps itself. | w-tf02790955, w-tf03512234, w-tf10072667 | open: three causes. tf10072667: wp:inline text boxes inside table cells are lifted out as body shapes anchored to the table (TextBox blocks with Inline=true, no X/Y) and all drawn at the table's top-left, so they overlap; they need laying out in their cell paragraph's line. tf03512234: the table-level w:tblCellMar top (5400 twips) is in the model (TableBlock.PaddingTop 360px) but not applied to the cells' content of the exact-height row. tf02790955: a w:titlePg section with no first-page header still reserves the default header's height on page 1, and a table style's conditional rPr (firstRow sz) is merged as paragraph-mark run properties (WordEditorPackageReader paragraphStyle) and so overrides the Title paragraph style; Word puts table style below a non-default paragraph style |
| W4 | Paragraph spacing is larger than Word's. | w-tf56150113, w-tf03458347, w-tf10002131 | fixed (cd98a6f6): an empty w:pPrDefault left after-spacing unset and the layout used its 10.7px default; Word uses 0 when w:pPrDefault is present (checked on winvm: without it Word does add space after). Line pitch is still ~1px taller than Word's in Normal paragraphs (Layout.fs line box factors) |
| W5 | A page background colour is missing. | w-tf33478217 | fixed (cd98a6f6): the page background is a full-page rectangle in a theme colour (accent1 lumMod/lumOff): same cause as W1 |
| W6 | A full-bleed picture starts about 90px too low. | w-tf22851674 | open: the header picture has a:xfrm rot="16200000" (270°); pictures (TextBuffer.ImageRun) have no rotation at all, so it is drawn unrotated, 1104x853 instead of 853x1104. Needs a Rotation field through DocxReader.ImageData, the editor package, ImageRun, the painter (DrawImage turned about its centre) and the writer (only flips are carried now) |
| W7 | Checkbox glyphs (Wingdings/MS Gothic ☐) draw as missing-glyph boxes. | w-tf56090272, docx4j-invoice-2013 | fixed (bullets: rbt-office cd98a6f6; checkboxes: redbat 0e784dfd4 measures missing-glyph runs as the browser draws them) |
| W8 | Comment anchors draw on the page as visible "[Comment 1]" boxes. | w-dfe-report | open: a comment reference is modelled as the visible text "[Comment N]" (WordEditorPackageReader, Document.commentMarkRanges, the comment-mark browser check) and laid out like any text. Word shows nothing inline. Needs the mark measured as zero width and not painted (Layout.fs text segments, hit testing, Paint), or a zero-width mark character in the model |
| W9 | On open the caret goes into the first text box, which shows its outline until you click elsewhere. Word puts the caret at the start of the body. | w-tf16392937, w-tf16392017 | fixed (cd98a6f6): it is not a text box but the first content control (the caret starts in it); w15:appearance="hidden" was ignored. Controls with hidden appearance get no frame or title tab (Document.contentControlFrameMemberships) |
| W10 | Run character spacing (`w:spacing w:val` in rPr, letter-spacing) isn't drawn: Word spreads "N O R T H W I N D", we draw normal spacing. | ~/Dev/office-hype/docs/designed/annual-review.docx (title band kicker) | fixed (redbat ab6aeeb42: the text stack measured and painted without it) |
| W11 | A table in a compatibilityMode ≤14 document starts at the margin; Word shifts it left by the first cell's left margin, so cell text lines up with body text. | annual-review.docx (the KPI table; python-docx's default template) | open: the compatibility mode is not read. The engine shifts a table left by its table-level default cell margin (WordLayout.Table.outerCellInsets: table/row w:tblCellMar, 0 when none), never by the first cell's own w:tcMar; Word (compat 14) shifts by the first cell's left margin (9.3px here). Needs compatibilityMode in the model and outerCellInsets to use the first cell's margin for compat <= 14 |
| W12 | Body text measures wider than Word's: on the same 650px column ours wraps a word earlier (Calibri 10.5pt, line spacing 1.15). | annual-review.docx ("…Ridge Park array a / quarter") | fixed (cd98a6f6): the widths match Word exactly (Segoe UI advances); the line broke early because a word's trailing space had to fit 3.5px inside the margin (LongTrailingSpaceFitSafety). With a host text measurer the editor now uses the PDF export's 0.25px |
| W13 | Drop caps (`w:framePr w:dropCap="drop" w:lines="3"`) aren't drawn as drop caps: the letter becomes a separate large paragraph above the text. | ~/Dev/office-hype/docs/designed/field-notes.docx | open: a section after a section break is laid out by the column flow (Layout.fs flowColumnsKept), which lays paragraphs out without obstacles and has no drop-cap handling; only layoutDocumentPassKept places a drop cap (pendingDropCap / placeWithDropCap). Needs the same in the column flow |
| W14 | Section columns (`w:cols w:num="2"`, continuous section) are laid out as one column. | field-notes.docx | fixed (cd98a6f6): the sectPr has two w:cols (space 720, then num 2); the reader took the first. It now takes the last, as Word does. (Ours balances the columns of the last section where Word fills the first column first: open) |
| W15 | A table whose `tblBorders` set every side (including insideH/insideV) to `nil` still shows thin light lines between cells, visible on a dark shaded header row; Word draws none. | annual-review.docx (sites table header) | fixed (cd98a6f6): not borders: adjacent cell fills at fractional pixel edges were each antialiased, leaving a light seam. Fills are now snapped to device pixels (Renderer snappedFill) |
| W16 | (editing) Typing over a selection that starts at the beginning of a paragraph loses the selected text's formatting: "4.8 GW" (Georgia 22pt teal, the paragraph's only run) replaced by typing becomes Normal (Calibri 10.5). Word gives typed text the formatting of the first selected character. | annual-review.docx: select "4.8 GW" in the first KPI cell, type | fixed (1636226b) |
| W18 | Content controls bound to custom XML (`w:dataBinding` with an xpath into a customXml part) show the text stored in the document; Word refreshes them from the bound XML on open ("John Funky", a "Cool!!" table where we show "Rich Word content can go here", "note 1"). | docx4j-invoice-2013 (~/Dev/office-hype/docs/test-invoice.docx) | open |

## Editor UI

| # | Difference | Status |
|---|---|---|
| U1 | (redbat) The Word toolbar's table tools are always present, disabled outside a table, so the toolbar wraps to two rows below about 1700px. Before, the toolbar wrapped and unwrapped as the caret entered tables, and the page jumped. Now a single "Table ▾" menu: one row, same height in and out of tables. | fixed (redbat 3567e0e2f) |
