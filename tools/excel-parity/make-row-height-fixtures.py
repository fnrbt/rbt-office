#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) for the row and column rules of a sheet as Excel 2019 opens it
(found on the real-estate models, each checked in Excel 2019 on the VM; tools/xlsx-browser check `rows-excel-parity`,
tests/office-formats/xlsx-bridge.fsx "rows:"). The helpers are make-sizing-fixtures.py's.

- synthetic-row-heights.xlsx (Calibri 11):
  - "Custom": rows of `ht` with `customHeight` from 15 to 16 and 19.5 to 20.5 pt by 0.05, and a few more: the
    pixels Excel gives a height of its own.
  - "Def157", "Def1595", "Def20", "Def16", "Def144": `sheetFormatPr customHeight="1"` defaults: the pixels of the
    sheet's own default.
  - "FitMin30": a default of 30 pt of its own and a row of a height of its own (so rows are fitted): a row of
    8 pt text, of 11 pt text, of 20 pt text - is a fitted row at least the default, or the font's line?
  - "FitMinCols": the same, a default of 16 pt of its own, every column in the style of Calibri 8.
  - "Fonts": a cell of 10.5 / 7.5 / 13.5 / 9.5 / 11.5 pt (Calibri, Arial, Times New Roman), each in a row.
- synthetic-hidden-axes.xlsx (Calibri 11):
  - "ColsToEnd": `<col min="4" max="16384" hidden="1">`: every column from D to XFD is hidden.
  - "DefColZero": `defaultColWidth="0"`: the columns of no `<col>` are hidden.
  - "ZeroHeight": `zeroHeight="1"`: the rows of no `<row>` are hidden - between the listed rows and past them.
  - "ZeroHeightCustom": the same with a default height of its own (15.95 pt) and rows of that height.
- synthetic-row-merged.xlsx (Calibri 11): "Merged": rows of merged cells - across columns (wrapped long text, wrapped
  short text, line breaks, 20 pt text, wrapped 20 pt text) and across rows (20 pt text over three rows); "ColStyle" /
  "ColStyleSmall": a column in the style of a larger / smaller font under wrapped text; "Nbsp": wrapped text of
  no-break spaces (Excel breaks a line at a space, not at a no-break space).
- synthetic-row-baseline.xlsx (Arial 10): a row's fonts meet at their baseline when aligned at the bottom - an empty
  Arial 12 cell and Arial 11 bold text, at the bottom, centred, at the top, mixed; wrapped Calibri 11 under an empty
  Calibri 20 cell, at the bottom, centred, at the top.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-row-height-fixtures.py
"""
import os
import runpy

HERE = os.path.dirname(os.path.abspath(__file__))
_helpers = runpy.run_path(os.path.join(HERE, 'make-sizing-fixtures.py'))
package, styles, sheet, inline, col = (_helpers[n] for n in ('package', 'styles', 'sheet', 'inline', 'col'))
OUTDIR = _helpers['OUTDIR']
CALIBRI_11 = _helpers['CALIBRI_11']
LONG = _helpers['LONG']


def font(face, size, extra=''):
    return f'{extra}<sz val="{size}"/><name val="{face}"/><family val="2"/>'


def xf(font_id, wrap=False):
    body = '<alignment wrapText="1"/>' if wrap else ''
    return (f'<xf numFmtId="0" fontId="{font_id}" fillId="0" borderId="0" xfId="0" applyFont="1"'
            + (' applyAlignment="1">' + body + '</xf>' if wrap else '/>'))


def text_row(r, style, text, attrs=''):
    return f'<row r="{r}"{attrs}>' + inline(f'A{r}', style, text) + '</row>'


# Fonts: 0 Calibri 11 (Normal), 1 Calibri 20, 2 Calibri 8, 3.. the fractional sizes.
FRACTIONS = [('Calibri', 10.5), ('Calibri', 7.5), ('Calibri', 13.5), ('Calibri', 9.5), ('Calibri', 11.5),
             ('Arial', 10.5), ('Arial', 7.5), ('Arial', 13.5), ('Arial', 11.5),
             ('Times New Roman', 10.5), ('Times New Roman', 13.5), ('Calibri', 15.5), ('Arial', 17.5)]
FONTS = [CALIBRI_11, font('Calibri', 20), font('Calibri', 8)] + [font(face, size) for face, size in FRACTIONS]
XFS = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>', xf(1), xf(2), xf(0, wrap=True)] \
    + [xf(3 + i) for i in range(len(FRACTIONS))] + [xf(1, wrap=True)]
S_BIG, S_SMALL, S_WRAP, S_FRACTION, S_BIG_WRAP = 1, 2, 3, 4, 4 + len(FRACTIONS)

CUSTOM = [round(15 + i * 0.05, 2) for i in range(21)] + [round(19.5 + i * 0.05, 2) for i in range(21)] \
    + [12, 12.3, 12.375, 12.5, 12.7, 14.25, 14.4, 15.6, 15.95, 0.75, 1, 2, 30.1, 30.4, 30.6, 30.9, 409]
custom_rows = [text_row(i + 1, 0, f'ht {h}', f' ht="{h}" customHeight="1"') for i, h in enumerate(CUSTOM)]


def default_sheet(height):
    return sheet([text_row(1, 0, f'default {height}'), text_row(3, 0, 'row 3')],
                 f'<sheetFormatPr defaultRowHeight="{height}" customHeight="1"/>')


fit_rows = [
    text_row(1, 0, 'custom 40', ' ht="40" customHeight="1"'),
    text_row(2, S_SMALL, 'Calibri 8, ht 12', ' ht="12"'),
    text_row(3, 0, 'Calibri 11'),
    text_row(4, S_BIG, 'Calibri 20'),
    text_row(5, S_SMALL, 'Calibri 8'),
    '<row r="6" ht="12"/>',
]
fraction_rows = [text_row(i + 1, S_FRACTION + i, f'{face} {size}') for i, (face, size) in enumerate(FRACTIONS)]

package(os.path.join(OUTDIR, 'synthetic-row-heights.xlsx'), [
    ('Custom', sheet(custom_rows, '<sheetFormatPr defaultRowHeight="15"/>', selected=True)),
    ('Def157', default_sheet(15.7)),
    ('Def1595', default_sheet(15.95)),
    ('Def20', default_sheet(20)),
    ('Def16', default_sheet(16)),
    ('Def144', default_sheet(14.4)),
    ('FitMin30', sheet(fit_rows, '<sheetFormatPr defaultRowHeight="30" customHeight="1"/>')),
    ('FitMinCols', sheet(fit_rows, '<sheetFormatPr defaultRowHeight="16" customHeight="1"/>', col(1, 16384, 9.140625, S_SMALL))),
    ('Fonts', sheet(fraction_rows, '<sheetFormatPr defaultRowHeight="15"/>')),
], styles(FONTS, xfs=XFS))

zero_rows = [
    text_row(1, 0, 'row 1'), text_row(2, 0, 'row 2'), text_row(4, 0, 'row 4 (3 not listed)'),
    text_row(5, 0, 'row 5 hidden', ' hidden="1"'), text_row(6, 0, 'row 6 ht 30', ' ht="30" customHeight="1"'),
    '<row r="7"/>', text_row(9, 0, 'row 9 ht 20', ' ht="20"'), text_row(10, S_BIG, 'row 10 Calibri 20'),
]
zero_custom_rows = [
    text_row(1, 0, 'row 1', ' ht="15.95" customHeight="1"'), text_row(2, 0, 'row 2', ' ht="15.95" customHeight="1"'),
    text_row(4, 0, 'row 4 no ht'), text_row(5, 0, 'row 5 hidden', ' ht="15.95" hidden="1" customHeight="1"'),
]
plain = [text_row(r, 0, f'row {r}') for r in (1, 2, 3)]
package(os.path.join(OUTDIR, 'synthetic-hidden-axes.xlsx'), [
    ('ColsToEnd', sheet(plain, '<sheetFormatPr defaultRowHeight="15"/>', col(4, 16384, 9.140625).replace('/>', ' hidden="1"/>'), selected=True)),
    ('DefColZero', sheet(plain, '<sheetFormatPr defaultColWidth="0" defaultRowHeight="15"/>', col(1, 3, 10))),
    ('ZeroHeight', sheet(zero_rows, '<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/>')),
    ('ZeroHeightCustom', sheet(zero_custom_rows, '<sheetFormatPr defaultColWidth="0" defaultRowHeight="15.95" customHeight="1" zeroHeight="1"/>',
                               col(1, 3, 10) + '<col min="4" max="16384" width="0" hidden="1"/>')),
], styles(FONTS, xfs=XFS))

# Merged cells. Column B 12 characters wide (89 px), C default.
MERGED_ROWS = [
    '<row r="1">' + inline('A1', S_WRAP, LONG) + '<c r="B1" s="3"/></row>',
    '<row r="2">' + inline('A2', S_WRAP, 'ok') + '<c r="B2" s="3"/></row>',
    '<row r="3">' + inline('A3', S_BIG, 'Big merged') + '<c r="B3" s="1"/></row>',
    '<row r="4">' + inline('A4', S_BIG, 'Big over 3 rows') + '</row>', '<row r="5"><c r="A5" s="1"/></row>', '<row r="6"><c r="A6" s="1"/></row>',
    '<row r="7">' + inline('A7', S_WRAP, 'one&#10;two&#10;three') + '<c r="B7" s="3"/></row>',
    '<row r="8">' + inline('A8', S_WRAP, LONG) + '<c r="B8" s="3"/><c r="C8" s="3"/></row>',
    '<row r="9">' + inline('A9', S_WRAP, LONG) + '<c r="B9" s="3"/>' + inline('D9', S_BIG, 'Big') + '</row>',
    '<row r="10">' + inline('A10', S_WRAP, LONG) + '</row>',
    '<row r="11">' + inline('A11', S_BIG_WRAP, 'Big wrapped merged text that is long') + '<c r="B11" s="3"/></row>',
]
MERGES = ['A1:B1', 'A2:B2', 'A3:B3', 'A4:A6', 'A7:B7', 'A8:C8', 'A9:B9', 'A11:B11']
# A column in the style of a larger font under wrapped text: column D of Calibri 20, F of Calibri 8.
COLSTYLE_ROWS = [
    '<row r="1">' + inline('B1', S_WRAP, LONG) + '</row>',
    '<row r="2">' + inline('B2', S_WRAP, 'one&#10;two&#10;three') + '</row>',
    '<row r="3">' + inline('B3', 0, 'plain') + '</row>',
    '<row r="4">' + inline('B4', S_WRAP, 'one&#10;two') + '</row>',
]
# Wrapped text of no-break spaces (U+00A0) in a column of 84 px: does Excel break a line at one?
WORDS = 'alpha beta gamma delta epsilon zeta eta theta'
NBSP_ROWS = [
    '<row r="1">' + inline('B1', S_WRAP, WORDS) + '</row>',
    '<row r="2">' + inline('B2', S_WRAP, WORDS.replace(' ', ' ')) + '</row>',
    '<row r="3">' + inline('B3', S_WRAP, 'alpha beta gamma delta epsilon zeta') + '</row>',
    '<row r="4">' + inline('B4', S_WRAP, 'alphabetagammadeltaepsilon') + '</row>',
]
package(os.path.join(OUTDIR, 'synthetic-row-merged.xlsx'), [
    ('Merged', sheet(MERGED_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12),
                     extra=f'<mergeCells count="{len(MERGES)}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in MERGES) + '</mergeCells>',
                     selected=True)),
    ('ColStyle', sheet(COLSTYLE_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12) + col(4, 4, 9.140625, S_BIG))),
    ('ColStyleSmall', sheet(COLSTYLE_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12) + col(4, 4, 9.140625, S_SMALL))),
    ('Nbsp', sheet(NBSP_ROWS, '<sheetFormatPr defaultRowHeight="15"/>', col(2, 2, 12))),
], styles(FONTS, xfs=XFS))

# A row's fonts at their baseline, by vertical alignment and with empty cells (vertex42's rows of an empty Arial 12
# cell and Arial 11 bold text, centred, in an Arial 10 sheet: 20 px in Excel). Workbook of Arial 10; fonts 1 Arial 12,
# 2 Arial 11 bold, 3 Calibri 20, 4 Calibri 11.
def aligned(font_id, vertical=None, wrap=False):
    body = ''.join([f' vertical="{vertical}"' if vertical else '', ' wrapText="1"' if wrap else ''])
    return (f'<xf numFmtId="0" fontId="{font_id}" fillId="0" borderId="0" xfId="0" applyFont="1"'
            + (f' applyAlignment="1"><alignment{body}/></xf>' if body else '/>'))


BASE_FONTS = [font('Arial', 10), font('Arial', 12), font('Arial', 11, '<b/>'), font('Calibri', 20), font('Calibri', 11)]
BASE_XFS = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            aligned(1), aligned(2), aligned(1, 'center'), aligned(2, 'center'), aligned(1, 'top'),  # 1..5
            aligned(4, wrap=True), aligned(4, 'center', wrap=True), aligned(4, 'top', wrap=True), aligned(3)]  # 6..9
A12, A11B, A12C, A11BC, A12T, C11W, C11WC, C11WT, C20 = range(1, 10)


def cells(r, *specs):
    out = ''
    for col, style, text in specs:
        out += f'<c r="{col}{r}" s="{style}"/>' if text is None else inline(f'{col}{r}', style, text)
    return f'<row r="{r}">' + out + '</row>'


BASE_ROWS = [
    cells(1, ('A', A12, None), ('B', A11B, 'bold 11')),            # empty Arial 12 + Arial 11 bold, bottom
    cells(2, ('A', A12, 'Arial 12'), ('B', A11B, 'bold 11')),       # both text, bottom
    cells(3, ('A', A12C, None), ('B', A11BC, 'bold 11')),          # vertex42's: empty, centred
    cells(4, ('A', A12C, 'Arial 12'), ('B', A11BC, 'bold 11')),     # both text, centred
    cells(5, ('A', A12C, 'Arial 12'), ('B', A11B, 'bold 11')),      # Arial 12 centred, bold bottom
    cells(6, ('A', A12T, 'Arial 12'), ('B', A11B, 'bold 11')),      # Arial 12 top, bold bottom
    cells(7, ('A', A12, 'Arial 12'), ('B', A11BC, 'bold 11')),      # Arial 12 bottom, bold centred
    cells(8, ('B', C11W, 'one&#10;two&#10;three'), ('D', C20, None)),   # 3 lines bottom + empty Calibri 20
    cells(9, ('B', C11WC, 'one&#10;two&#10;three'), ('D', C20, None)),  # ... centred
    cells(10, ('B', C11WT, 'one&#10;two&#10;three'), ('D', C20, None)), # ... top
    cells(11, ('B', C11W, 'one&#10;two&#10;three'), ('D', C20, 'Big')), # ... with text
]
package(os.path.join(OUTDIR, 'synthetic-row-baseline.xlsx'), [
    ('Baseline', sheet(BASE_ROWS, '<sheetFormatPr defaultRowHeight="12.75"/>', selected=True)),
], styles(BASE_FONTS, xfs=BASE_XFS))
