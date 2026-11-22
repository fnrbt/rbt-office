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
          colstyles=None):
    return dict(name=name, cells=cells, cols=cols or {}, heights=heights or {}, setup=setup, options=options,
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
          'widths-aptosnarrow11': ('widths', 'Aptos Narrow', 11)}
NORMAL.update({f'rowunit-{f.lower().replace(" ", "")}{z}': ('rowunit', f, z) for f, z in ROWUNIT_FONTS})


def build(out, which):
    fonts = FONTS
    if which in NORMAL:
        which, face, size = NORMAL[which]
        fonts = [f'<sz val="{size}"/><name val="{face}"/><family val="2"/>', f'<b/><sz val="{size}"/><name val="{face}"/><family val="2"/>']
    sheets = sets()[which]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(FILLS)}">' + ''.join(FILLS) + '</fills>'
        f'<borders count="{len(BORDERS)}">' + ''.join(BORDERS) + '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(XFS)}">' + ''.join(XFS) + '</cellXfs>'
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
        colxml = [f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>' for c, w in sorted(sh['cols'].items())]
        colxml += [f'<col min="{c}" max="{c}" width="9.140625" style="{st}"/>' for c, st in sorted(sh['colstyles'].items())]
        if colxml:
            cols = '<cols>' + ''.join(colxml) + '</cols>'
        merges = ''
        if sh['merges']:
            merges = f'<mergeCells count="{len(sh["merges"])}">' + ''.join(f'<mergeCell ref="{m}"/>' for m in sh['merges']) + '</mergeCells>'
        parts.append(DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">{sh["fmt"]}<sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{"".join(xml_rows)}</sheetData>{merges}'
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
