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

Status column: open / fixed (commit) / not a bug (why). Some examples were matched to
the report after the fact: if one doesn't show the difference, search the side-by-sides.

## Spreadsheet

| # | Difference | Examples | Status |
|---|---|---|---|
| S1 | Cell text is drawn one CSS px per point (11pt → 11px; Excel draws 14.67px), so all sheet text is about 25% small. Column widths and row heights match. | every workbook | open (in progress) |
| S2 | A cell's font family is ignored when the face isn't installed: Georgia, Constantia, Segoe UI, Calibri, Century Gothic and Franklin Gothic all fall back to a generic sans. The Word side ships web fonts for some of these. | x-tf03986924, x-tf04035471 | open |
| S3 | Sparklines aren't drawn. | x-tf02802361 | open |
| S4 | Pattern (hatched) fills aren't drawn. | x-tf03107663 | open |
| S5 | Rich-text runs in a cell lose their colours ("OPEN vs. CLOSED", "Step one:"). | x-tf03784001 | open |
| S6 | Accounting format doesn't pin `$` to the cell's left edge. | x-tf03458075 | open |
| S7 | Icon sets look different from Excel's. | x-tf02897386 | open |
| S8 | `###` where Excel shows the number (attendance sheet). | x-tf02780235 | open |
| S9 | A pictogram shows its alt text ("decorative element") instead of the picture. | x-tf02930056 | open |
| S10 | Chart series labels spill outside the chart; pie data labels are missing. | x-tf03783862, x-tf03986584 | open |
| S11 | Filter buttons show on table headers where Excel hides them (table with `showAutoFilter`/header buttons off). | x-tf03427338 | open |
| S12 | A General number too wide for its cell shows `4E+...`; Excel shrinks it to `4E+05`. | `=SUM` of salaries in a narrow column | open |

## Word

| # | Difference | Examples | Status |
|---|---|---|---|
| W1 | Shapes in groups or anchored to the page often go missing: coloured panels, photos inside groups, logo blocks, background art. | w-tf16412190, w-tf16392850, w-tf16392545, w-tf56355079, w-tf04011831, w-tf77989015, w-tf56090272 | open |
| W2 | A group's fill or outline loses its colour. | w-tf02933929 (candles), w-tf22897213, w-tf00241780 | open |
| W3 | Text boxes are placed wrongly; one brochure's text overlaps itself. | w-tf02790955, w-tf03512234, w-tf10072667 | open |
| W4 | Paragraph spacing is larger than Word's. | w-tf56150113, w-tf03458347, w-tf10002131 | open |
| W5 | A page background colour is missing. | w-tf33478217 | open |
| W6 | A full-bleed picture starts about 90px too low. | w-tf22851674 | open |
| W7 | Checkbox glyphs (Wingdings/MS Gothic ☐) draw as missing-glyph boxes. | w-tf56090272, docx4j-invoice-2013 | open |
| W8 | Comment anchors draw on the page as visible "[Comment 1]" boxes. | w-dfe-report | open |
| W9 | On open the caret goes into the first text box, which shows its outline until you click elsewhere. Word puts the caret at the start of the body. | w-tf16392937, w-tf16392017 | open |

## Editor UI

| # | Difference | Status |
|---|---|---|
| U1 | (redbat) The Word toolbar's table tools are always present, disabled outside a table, so the toolbar wraps to two rows below about 1700px. Before, the toolbar wrapped and unwrapped as the caret entered tables, and the page jumped. Now a single "Table ▾" menu: one row, same height in and out of tables. | fixed (redbat 3567e0e2f) |
