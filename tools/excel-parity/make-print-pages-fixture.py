#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) for how Excel 2019 lays a sheet out on pages when it exports it to PDF:
which cells it prints without a print area, where it breaks pages, what scale fit-to-page comes to, how high a row
prints (docs/excel-parity-tracker.md, "Print pages"). Checked with `excel-print-pages.sh` on the VM (Page Break
Preview's breaks and each sheet's PDF) and `compare-print-pages.fsx` / `pdftotext -bbox` here.

Sheets (SHEETS below): where the printed range starts and ends (values only from C5; a fill, a border, a bold font,
a number format past the values), rows of given heights, columns fitted to a page wide, explicit scales, long sheets
broken automatically.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-print-pages-fixture.py OUT.xlsx [set]
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

# Styles: 0 Normal, 1 a yellow fill, 2 a thin box border, 3 bold, 4 top aligned, 5 a number format (0.00), 6 bold top aligned
FONTS = ['<sz val="11"/><name val="Calibri"/><family val="2"/>', '<b/><sz val="11"/><name val="Calibri"/><family val="2"/>']
FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>',
         '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>']
BORDERS = ['<border><left/><right/><top/><bottom/><diagonal/></border>',
           '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right>'
           '<top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>']
XFS = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
       '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
       '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>',
       '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
       '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>',
       '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
       '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top"/></xf>']
# Style 7 (a date, d-mmm-yy), only in the sets that use it (the others' bytes as they were probed).
DATE_XF = '<xf numFmtId="15" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
# Styles 8 - 13 of the bestfit-figures sets: #,##0; "$"#,##0; #,##0_);(#,##0); 0.0%; 0.00; d-mmm-yy.
FIGURE_NUMFMTS = ('<numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/>'
                  '<numFmt numFmtId="165" formatCode="#,##0_);\\(#,##0\\)"/><numFmt numFmtId="166" formatCode="0.0%"/></numFmts>')
FIGURE_XFS = [f'<xf numFmtId="{n}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' for n in (3, 164, 165, 166, 2, 15)]
# Styles 8 on of the bestfit-chars sets: a number with literal characters after it (which characters an AutoFit
# column's print width counts, and how wide), percentages, and fonts other than the Normal one (2: Calibri 20, 3: Arial 10,
# 1: the Normal font bold).
CHAR_LITS = ['i', 'ii', 'iiii', 'iiiiiiii', 'W', 'WW', 'WWWW', 'WWWWWWWW', '.', '....', '........', ',,,,', '%%%%',
             '    ', '0000', 'MMMM', '----']
CHAR_FMTS = ['0'] + [f'0"{t}"' for t in CHAR_LITS] + ['0.0%', '0.00', '#,##0', '0%', '0.000%', '"$"#,##0_)']
CHAR_FONTS = ['<sz val="20"/><name val="Calibri"/><family val="2"/>', '<sz val="10"/><name val="Arial"/><family val="2"/>']


BORDER_STYLES = ['hair', 'thin', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'mediumDashDotDot', 'slantDashDot',
                 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double']


def char_style(fmt, font=0):
    """The style index (bestfit-chars sets) of a number format of CHAR_FMTS in a font (0 Normal, 2, 3, 1 as above)."""
    return 8 + CHAR_FMTS.index(fmt) + len(CHAR_FMTS) * [0, 2, 3, 1].index(font)


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def block(r0, c0, r1, c1, style=0):
    """Labels r<row>c<col> in every cell of the block (1-based, inclusive)."""
    return [(r, c, f'r{r}c{c}', style) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)]


def sheet(name, cells, cols=None, heights=None, setup='<pageSetup paperSize="1" orientation="portrait"/>',
          options='<printOptions gridLines="1"/>', margins=None, breaks='', names=None, fmt='', merges=None, rowstyles=None,
          colstyles=None, hidden=None, autofilter=''):
    return dict(name=name, cells=cells, cols=cols or {}, heights=heights or {}, setup=setup, options=options,
                hidden=hidden or set(), autofilter=autofilter,
                merges=merges or [], rowstyles=rowstyles or {}, colstyles=colstyles or {},
                margins=margins or '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
                breaks=breaks, names=names or [], fmt=fmt)


FIT = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'


def sets():
    out = {}
    # Where the printed range is, without a print area.
    out['range'] = [
        sheet('StartC5', block(5, 3, 12, 6)),
        sheet('FillPast', block(1, 1, 5, 3) + [(12, 6, None, 1)]),
        sheet('BorderPast', block(1, 1, 5, 3) + [(12, 6, None, 2)]),
        sheet('BoldPast', block(1, 1, 5, 3) + [(12, 6, None, 3)]),
        sheet('FormatPast', block(1, 1, 5, 3) + [(12, 6, None, 5)]),
        sheet('StartC5Wide', block(5, 3, 60, 14)),
        sheet('Empty', []),
        # Text running past the last column with something in it.
        sheet('Overflow', block(1, 1, 5, 3) + [(7, 3, 'The quick brown fox jumps over the lazy dog', 0)]),
        sheet('OverflowFar', block(1, 1, 5, 3) + [(7, 3, ' '.join(['The quick brown fox jumps over the lazy dog'] * 6), 0)]),
        sheet('OverflowBlocked', block(1, 1, 5, 3) + [(7, 3, 'The quick brown fox jumps over the lazy dog', 0), (7, 5, None, 1)]),
        sheet('FormulaBlank', block(1, 1, 5, 3) + [(12, 6, ('=""', ''), 0)]),
        sheet('MergePast', block(1, 1, 5, 3) + [(10, 5, 'm', 0)], merges=['E10:H14']),
        sheet('MergeFill', block(1, 1, 5, 3) + [(10, 5, None, 1)], merges=['E10:H14']),
        sheet('RowFill', block(1, 1, 5, 3), rowstyles={12: 1}),
        sheet('ColFill', block(1, 1, 5, 3), colstyles={8: 1}),
        sheet('OverflowNumber', block(1, 1, 5, 3) + [(7, 3, 12345678901234567890, 0)]),
    ]
    # How high a row prints: top-aligned labels in rows of given heights (pt), gridlines printed.
    hs = [15, 7.5, 10, 12, 12.75, 14.25, 15.75, 16.5, 18, 20, 24, 30, 45, 60, 75, 99.75, 150, 200, 300, 409.5]
    out['heights'] = [
        sheet('Heights', [(i + 1, 1, f'h{h}', 4) for i, h in enumerate(hs)], heights={i + 1: h for i, h in enumerate(hs)}),
        sheet('Default', [(r, 1, f'r{r}', 4) for r in range(1, 121)]),
        sheet('Bold', [(r, 1, f'r{r}', 6) for r in range(1, 121)]),
    ]
    # Fit to a page wide / tall, and explicit scales: the breaks and the scale (the PDF's column pitch).
    out['fit'] = [
        sheet(f'Fit{w}', block(1, 1, 40, w), setup='<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>', fmt=FIT)
        for w in (10, 11, 13, 17, 23)
    ] + [
        sheet('FitTall', block(1, 1, 130, 4), setup='<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="2"/>', fmt=FIT),
        sheet('Fit2x2', block(1, 1, 140, 20), setup='<pageSetup paperSize="1" orientation="portrait" fitToWidth="2" fitToHeight="2"/>', fmt=FIT),
    ] + [
        sheet(f'Scale{s}', block(1, 1, 120, 12), setup=f'<pageSetup paperSize="1" orientation="portrait" scale="{s}"/>')
        for s in (55, 77, 90, 133)
    ]
    # Long sheets broken automatically: default rows, Letter / A4 / landscape, other margins.
    out['breaks'] = [
        sheet('Letter', block(1, 1, 200, 12)),
        sheet('A4', block(1, 1, 200, 12), setup='<pageSetup paperSize="9" orientation="portrait"/>'),
        sheet('Landscape', block(1, 1, 200, 16), setup='<pageSetup paperSize="1" orientation="landscape"/>'),
        sheet('Narrow', block(1, 1, 200, 12), margins='<pageMargins left="0.25" right="0.25" top="0.25" bottom="0.25" header="0" footer="0"/>'),
        sheet('Wide', block(1, 1, 200, 12), margins='<pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>'),
    ]
    # How wide a column prints: a left-aligned "x" in row 1 of columns of these widths (characters), A to AN.
    ws = [0.5, 0.75, 1, 1.25, 1.5, 1.664, 2, 2.5, 3, 4, 5, 6, 7, 8, 8.43, 9, 9.109375, 10, 11, 12, 12.89, 13.55, 15, 18.11,
          20, 25, 1, 2, 3, 0.25, 0.1, 1.1, 1.33, 1.9, 2.2, 2.7]
    out['widths'] = [
        sheet('Widths', [(1, c + 1, 'x', 0) for c in range(len(ws))] + [(2, c + 1, w, 0) for c, w in enumerate(ws)],
              cols={c + 1: w for c, w in enumerate(ws)}, setup='<pageSetup paperSize="1" orientation="landscape"/>',
              margins='<pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'),
    ]
    # Rows / columns to repeat inside and above the print area, and a print area written as a formula.
    out['titles'] = [
        sheet('TitlesInside', block(1, 1, 200, 6), names=[('_xlnm.Print_Area', "OFFSET(TitlesInside!$A$1,0,0,100+TitlesInside!$H$1,6)"),
                                                          ('_xlnm.Print_Titles', 'TitlesInside!$20:$20')]
              ),
        sheet('TitlesTop', block(1, 1, 200, 6), names=[('_xlnm.Print_Titles', 'TitlesTop!$1:$2')]),
        sheet('TitlesAbove', block(1, 1, 200, 6), names=[('_xlnm.Print_Area', 'TitlesAbove!$A$10:$F$200'), ('_xlnm.Print_Titles', 'TitlesAbove!$1:$1')]),
        sheet('TitlesCols', block(1, 1, 30, 30), names=[('_xlnm.Print_Titles', 'TitlesCols!$A:$B')]),
        sheet('TitlesColsInside', block(1, 1, 30, 30), names=[('_xlnm.Print_Titles', 'TitlesColsInside!$E:$E')]),
    ]
    out['titles'][0]['cells'].append((1, 8, 37, 0))
    # Columns Excel's AutoFit left (`bestFit`) next to ones it did not: a number too wide for its column, text, a date.
    bf = {2: (5, 'bestFit'), 3: (5, 'bestFit'), 4: 5, 5: (20, 'bestFit'), 6: (9.5546875, 'bestFit'), 7: 9.5546875, 8: (5, 'bestFit'), 9: 5}
    out['bestfit'] = [
        sheet('BestFit', [(1, c, 'x', 0) for c in range(1, 11)] + [(2, 2, 123456789, 0), (2, 3, 'Long text here', 0), (2, 4, 123456789, 0),
                                                                    (2, 5, 1, 0), (2, 6, 44317, 7), (2, 7, 44317, 7), (2, 8, 1.5, 0), (2, 9, 1.5, 0)],
              cols=bf),
    ]
    # How wide a column Excel's AutoFit left (`bestFit`) prints: such columns of widths 1 to 25 with nothing in them,
    # a text, a number; numbers of 1 to 14 figures in narrow and default ones; dates and formats; bold; each set
    # beside one of the same widths without the flag. An "x" in A and past the last column.
    bw = [1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 20, 25]
    land = dict(setup='<pageSetup paperSize="1" orientation="landscape"/>',
                margins='<pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')

    def swept(name, widths, content, flag=True):
        cells = [(1, 1, 'x', 0), (1, len(widths) + 2, 'x', 0)]
        for i, w in enumerate(widths):
            v = content(i, w)
            if v is not None:
                cells.append((2, i + 2) + (v if isinstance(v, tuple) else (v, 0)))
        return sheet(name, cells, cols={i + 2: ((w, 'bestFit') if flag else w) for i, w in enumerate(widths)}, **land)
    figures = '123456789012345'
    out['bestfit-sweep'] = [
        swept('Empty', bw, lambda i, w: None),
        swept('EmptyPlain', bw, lambda i, w: None, flag=False),
        swept('Text', bw, lambda i, w: 'x'),
        swept('TextLong', bw, lambda i, w: 'The quick brown fox'),
        swept('One', bw, lambda i, w: 1),
        swept('OnePlain', bw, lambda i, w: 1, flag=False),
        swept('Figures3', [3] * 14, lambda i, w: int(figures[:i + 1])),
        swept('Figures8', [8.43] * 14, lambda i, w: int(figures[:i + 1])),
        swept('Figures3Plain', [3] * 14, lambda i, w: int(figures[:i + 1]), flag=False),
        swept('Fixed2', [3, 5, 8, 12] * 3, lambda i, w: (1234.5 * 10 ** (i // 4), 5)),
        swept('Dates', [3, 5, 8, 9.5546875, 12], lambda i, w: (44317, 7)),
        swept('Bold', bw, lambda i, w: ('x', 3)),
        swept('Halves', [4.5, 5.5, 6.25, 7.75, 8.86, 9.14, 10.71, 13.29], lambda i, w: 'x'),
    ]
    # The width a number needs in a column AutoFit left: columns of one character, each with a General number of 1 to
    # 11 figures; numbers in formats (8 - 13: #,##0, "$"#,##0, #,##0_);(#,##0) of a negative, 0.0%, 0.00, d-mmm-yy)
    # and letters as text. For `bestfit-figures-<face><size>` (the Normal font).
    fmts = [(1234567, 8), (12345, 9), (-1234567, 10), (0.125, 11), (1.5, 12), (44317, 13), (-12, 10), (999, 8),
            (1234567890, 8), (0.5, 11), (100, 9)]
    out['bestfit-figures'] = [
        swept('General', [1] * 11, lambda i, w: int(figures[:i + 1])),
        swept('Formats', [1] * len(fmts), lambda i, w: fmts[i]),
        swept('Wide', [12] * 6, lambda i, w: [1, 12345, 'x', (1234567, 8), (0.125, 11), 99999999][i]),
        swept('Letters', [1] * 8, lambda i, w: int('1' * (i + 1)) if i < 4 else int('8' * (i - 3))),
    ]
    # Which characters of a number's text the print width of a column AutoFit left counts, and by how much: a 1 with
    # literal characters after it, percentages of other values, figures of other digits, other fonts, wide columns.
    cs = char_style
    lit = lambda ts: [(1, cs('0'))] + [(1, cs(f'0"{t}"')) for t in ts]
    out['bestfit-chars'] = [
        swept('LitI', [1] * 9, lambda i, w: lit(['i', 'ii', 'iiii', 'iiiiiiii', 'W', 'WW', 'WWWW', 'WWWWWWWW'])[i]),
        swept('LitDot', [1] * 10, lambda i, w: lit(['.', '....', '........', ',,,,', '%%%%', '    ', '0000', 'MMMM', '----'])[i]),
        swept('Pct', [1] * 10, lambda i, w: [(0.125, cs('0.0%')), (0.5, cs('0.0%')), (0.111, cs('0.0%')), (0.999, cs('0.0%')),
                                             (0, cs('0.0%')), (0.05, cs('0.0%')), (0.5, 0), (0.125, 0), (0.5, cs('0%')), (0.5, cs('0.000%'))][i]),
        swept('Digits', [1] * 10, lambda i, w: [(v, cs('0')) for v in (1, 5, 7, 9, 11111, 55555, 99999, 10000, 44444, 77777)][i]),
        swept('Decimals', [1] * 8, lambda i, w: [(1.5, cs('0.00')), (5.5, cs('0.00')), (1.25, 0), (5.25, 0), (1234567, cs('#,##0')),
                                                  (5555555, cs('#,##0')), (1234567, cs('"$"#,##0_)')), (5555555, cs('"$"#,##0_)'))][i]),
        swept('Fonts', [1] * 9, lambda i, w: [(12345, cs('0')), (12345, cs('0', 2)), (12345, cs('0', 3)), (12345, cs('0', 1)),
                                               (1234567, cs('"$"#,##0_)', 1)), (1, cs('0', 2)), (1, cs('0', 3)), (1, cs('0"WWWW"', 2)),
                                               (1, cs('0"WWWW"', 3))][i]),
        swept('WideCols', [8, 10, 12, 15, 8, 10, 12, 15], lambda i, w: ([(1, cs('0'))] * 4 + [(1234567, cs('"$"#,##0_)'))] * 4)[i]),
        swept('TwoCells', [1] * 4, lambda i, w: (12345, cs('0'))),
    ]
    # (TwoCells: a second, longer number below two of them, a long text below another.)
    out['bestfit-chars'][-1]['cells'] += [(3, c, 1234567, cs('0')) for c in (3, 5)] + [(3, 4, 'a long text in the column', 0)]
    # Every border style in print: a box of each round cells B:F of its row (rows 2, 4, ... in BORDER_STYLES' order),
    # no gridlines; a row of each style 4 columns long under a cell's bottom only (rows 30, 32, ...).
    out['borders'] = [
        sheet('Borders', [(1, 1, 'x', 0)] + [(2 + 2 * i, c, None, 8 + i) for i in range(len(BORDER_STYLES)) for c in range(2, 7)]
              + [(30 + 2 * i, c, None, 8 + len(BORDER_STYLES) + i) for i in range(len(BORDER_STYLES)) for c in range(2, 6)]
              + [(2 + 2 * i, 8, s, 0) for i, s in enumerate(BORDER_STYLES)], options=''),
    ]
    # A list validation's cells (Excel shows its arrow beside the selected cell only): what prints of them. (The
    # <dataValidations> goes where the autoFilter would: there are no merged cells after it.)
    dv = ('<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
          'sqref="B2:B6 D2"><formula1>"Yes,No,Maybe"</formula1></dataValidation></dataValidations>')
    out['validation'] = [
        sheet('Validation', [(1, 1, 'x', 0)] + [(r, 2, 'Yes', 0) for r in range(2, 5)] + [(2, 4, 'No', 2)], autofilter=dv),
    ]
    # Rows hidden in an AutoFilter's range by a filter of no criteria (its buttons hidden, as houseflipping's REPAIR
    # ESTIMATOR): fit to a page wide and 6 tall, and broken by height.
    hide = set(range(12, 181))
    af = ('<autoFilter ref="A1:D200"><filterColumn colId="0" showButton="0"/><filterColumn colId="2" showButton="0"/>'
          '</autoFilter>')
    out['filter'] = [
        sheet('FilterFit', block(1, 1, 200, 4), hidden=hide, autofilter=af, fmt=FIT,
              setup='<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="6"/>'),
        sheet('FilterBreaks', block(1, 1, 200, 4), hidden=hide, autofilter=af),
        sheet('HiddenBreaks', block(1, 1, 200, 4), hidden=hide),
    ]
    # A row of the Normal font as it prints: 60 rows of the default height (and gridlines), for `rowunit-<face><size>`.
    out['rowunit'] = [sheet('Rows', [(r, 1, f'r{r}', 0) for r in range(1, 61)])]
    return out


# Normal fonts the row unit is measured for (`rowunit-tahoma10` ...).
ROWUNIT_FONTS = [('Tahoma', 8), ('Tahoma', 9), ('Tahoma', 10), ('Tahoma', 11), ('Verdana', 10), ('Segoe UI', 9),
                 ('Segoe UI', 10), ('Segoe UI', 11), ('Times New Roman', 10), ('Times New Roman', 12), ('Arial', 9),
                 ('Arial', 11), ('Calibri', 9), ('Calibri', 14), ('Aptos Narrow', 10), ('Aptos Narrow', 12),
                 ('Aptos', 12), ('Century Gothic', 10), ('Garamond', 11), ('Arial Narrow', 10), ('Calibri Light', 11),
                 ('Georgia', 10), ('Trebuchet MS', 10), ('Book Antiqua', 10)]


# Sets whose Normal style is another font: (set, Normal font name, size).
NORMAL = {'widths-arial12': ('widths', 'Arial', 12), 'widths-calibri12': ('widths', 'Calibri', 12),
          'widths-arial10': ('widths', 'Arial', 10), 'widths-arialmt12': ('widths', 'Arial MT', 12),
          'heights-arial10': ('heights', 'Arial', 10), 'heights-calibri12': ('heights', 'Calibri', 12),
          'heights-arial12': ('heights', 'Arial', 12), 'heights-cambria14': ('heights', 'Cambria', 14),
          'heights-aptosnarrow11': ('heights', 'Aptos Narrow', 11), 'heights-aptos11': ('heights', 'Aptos', 11),
          'heights-arial8': ('heights', 'Arial', 8), 'heights-calibri10': ('heights', 'Calibri', 10),
          'widths-aptosnarrow11': ('widths', 'Aptos Narrow', 11),
          'bestfit-sweep-arial10': ('bestfit-sweep', 'Arial', 10), 'bestfit-sweep-calibri14': ('bestfit-sweep', 'Calibri', 14)}
BESTFIT_FONTS = [('Calibri', 11), ('Calibri', 9), ('Calibri', 10), ('Calibri', 12), ('Calibri', 16), ('Calibri', 20),
                 ('Arial', 8), ('Arial', 10), ('Arial', 12), ('Arial', 14), ('Times New Roman', 12), ('Verdana', 10),
                 ('Segoe UI Black', 11), ('Aptos Narrow', 11), ('Tahoma', 10)]
NORMAL.update({f'bestfit-figures-{f.lower().replace(" ", "")}{z}': ('bestfit-figures', f, z) for f, z in BESTFIT_FONTS})
NORMAL.update({f'bestfit-chars-{f.lower().replace(" ", "")}{z}': ('bestfit-chars', f, z) for f, z in
               [('Calibri', 11), ('Calibri', 20), ('Arial', 10), ('Arial', 8), ('Calibri', 12)]})
NORMAL.update({f'rowunit-{f.lower().replace(" ", "")}{z}': ('rowunit', f, z) for f, z in ROWUNIT_FONTS})


def build(out, which):
    fonts = FONTS
    if which in NORMAL:
        which, face, size = NORMAL[which]
        fonts = [f'<sz val="{size}"/><name val="{face}"/><family val="2"/>', f'<b/><sz val="{size}"/><name val="{face}"/><family val="2"/>']
    sheets = sets()[which]
    xfs = XFS + ([DATE_XF] if which.startswith('bestfit') or which == 'borders' else []) + (FIGURE_XFS if which.startswith('bestfit-figures') else [])
    numfmts = FIGURE_NUMFMTS if which.startswith('bestfit-figures') else ''
    if which == 'bestfit-chars':
        fonts = fonts + CHAR_FONTS
        builtin = {'0': 1, '0.00': 2, '#,##0': 3, '0%': 9, '0.00%': 10}
        ids = {}
        custom = []
        for f in CHAR_FMTS:
            if f in builtin:
                ids[f] = builtin[f]
            else:
                ids[f] = 164 + len(custom)
                custom.append(f'<numFmt numFmtId="{ids[f]}" formatCode="{escape(f, {chr(34): "&quot;"})}"/>')
        numfmts = f'<numFmts count="{len(custom)}">' + ''.join(custom) + '</numFmts>'
        xfs = xfs + [f'<xf numFmtId="{ids[f]}" fontId="{font}" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
                     for font in (0, 2, 3, 1) for f in CHAR_FMTS]
    borders = BORDERS
    if which == 'borders':
        box = lambda st: f'<border><left style="{st}"><color indexed="64"/></left><right style="{st}"><color indexed="64"/></right><top style="{st}"><color indexed="64"/></top><bottom style="{st}"><color indexed="64"/></bottom><diagonal/></border>'
        under = lambda st: f'<border><left/><right/><top/><bottom style="{st}"><color indexed="64"/></bottom><diagonal/></border>'
        borders = BORDERS + [box(st) for st in BORDER_STYLES] + [under(st) for st in BORDER_STYLES]
        xfs = xfs + [f'<xf numFmtId="0" fontId="0" fillId="0" borderId="{2 + i}" xfId="0" applyBorder="1"/>' for i in range(2 * len(BORDER_STYLES))]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">' + numfmts +
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(FILLS)}">' + ''.join(FILLS) + '</fills>'
        f'<borders count="{len(borders)}">' + ''.join(borders) + '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    parts = []
    for i, sh in enumerate(sheets):
        rows = {}
        for r, c, v, st in sh['cells']:
            rows.setdefault(r, []).append((c, v, st))
        for r in list(sh['heights']) + list(sh['rowstyles']):
            rows.setdefault(r, [])
        xml_rows = []
        for r in sorted(rows):
            ht = f' ht="{sh["heights"][r]}" customHeight="1"' if r in sh['heights'] else ''
            if r in sh.get('hidden', ()):
                ht += ' hidden="1"'
            if r in sh['rowstyles']:
                ht += f' s="{sh["rowstyles"][r]}" customFormat="1"'
            cs = []
            for c, v, st in sorted(rows[r], key=lambda x: x[0]):
                ref = f'{col_letter(c)}{r}'
                sa = f' s="{st}"' if st else ''
                if v is None:
                    cs.append(f'<c r="{ref}"{sa}/>')
                elif isinstance(v, tuple):
                    cs.append(f'<c r="{ref}"{sa} t="str"><f>{escape(v[0][1:])}</f><v>{escape(v[1])}</v></c>')
                elif isinstance(v, str):
                    cs.append(f'<c r="{ref}"{sa} t="inlineStr"><is><t>{escape(v)}</t></is></c>')
                else:
                    cs.append(f'<c r="{ref}"{sa}><v>{v}</v></c>')
            xml_rows.append(f'<row r="{r}"{ht}>' + ''.join(cs) + '</row>')
        cols = ''
        # (A width given as (width, 'bestFit') is marked as Excel's AutoFit leaves it.)
        colxml = [f'<col min="{c}" max="{c}" width="{w[0] if isinstance(w, tuple) else w}"{" bestFit=\"1\"" if isinstance(w, tuple) else ""} customWidth="1"/>'
                  for c, w in sorted(sh['cols'].items())]
        colxml += [f'<col min="{c}" max="{c}" width="9.140625" style="{st}"/>' for c, st in sorted(sh['colstyles'].items())]
        if colxml:
            cols = '<cols>' + ''.join(colxml) + '</cols>'
        merges = ''
        if sh['merges']:
            merges = f'<mergeCells count="{len(sh["merges"])}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in sh['merges']) + '</mergeCells>'
        parts.append(DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{sh["fmt"]}<sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{"".join(xml_rows)}</sheetData>{sh.get("autofilter", "")}{merges}'
            f'{sh["options"]}{sh["margins"]}{sh["setup"]}{sh["breaks"]}</worksheet>'))
    names = ''.join(f'<definedName name="{n}" localSheetId="{i}">{escape(v)}</definedName>'
                    for i, sh in enumerate(sheets) for n, v in sh['names'])
    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>' for i in range(len(sheets)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                                   + ''.join(f'<sheet name="{sh["name"]}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, sh in enumerate(sheets))
                                   + '</sheets>' + (f'<definedNames>{names}</definedNames>' if names else '') + '</workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, xml in enumerate(parts):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(out)


if __name__ == '__main__':
    if len(sys.argv) > 2:
        build(sys.argv[1], sys.argv[2])
    else:
        for k in list(sets()) + list(NORMAL):
            build(sys.argv[1].replace('.xlsx', f'-{k}.xlsx'), k)
