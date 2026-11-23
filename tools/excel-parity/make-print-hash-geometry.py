#!/usr/bin/env python3
"""Writes a synthetic workbook for how Excel 2019 prints hash overflow (`####`) as the column gets wider.

Hypothesis under test (docs/excel-parity-tracker.md): for a fixed number of printed hashes, the word box's right
edge relative to the cell's LEFT physical edge advances by the change in cell physical width (slope 1) when the
cell is right-aligned, by half that (0.5) when centred, and not at all (0) when left-aligned. The hash run itself
is rigid, so run width = n * advance('#') at the printed face/size/scale.

Each sheet is one case (face, size, italic, alignment, print scale, overflow vs literal hash text). Every test
column in the sheet has a different character width and its own solid fill colour; the fill rectangle in the PDF
is the column's physical cell boundary, and the colour maps the printed word back to its width. Row 1 is a 13-digit numeric value (format `0`), which Excel prints as hashes unless the final 21-character
diagnostic column is wide enough for the whole number (the 'full-numeric' diagnostic). Row 2 is a non-overflow control
(`-98.41`, format `0.00`) for the italic-vs-upright translation. Literal sheets instead print the strings
`#######` (7) and `########` (8) as text, so the hash count is fixed by the string, not by the column width.

Widths are in Excel column-width characters. The 7->8 hash boundary for Aptos Narrow 14 / 100 % falls near 10.5,
so the scans bracket it on both sides. The 75 % scale sheet keeps the same column widths: the whole page scales,
so the hash count as a function of width characters must be unchanged while run width and both edge slopes scale
by 0.75 in PDF points.

A companion `.meta.tsv` records, per sheet column, the case and the width characters, so the measurement script
needs no guesswork. The na"ive `make-print-right-edge-fixture.py` remains the oracle for the last-glyph reserve;
this one owns width geometry and the hash-count transition.

    python3 tools/excel-parity/make-print-hash-geometry.py OUT.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

WIDTHS_APTOS = [round(8.0 + 0.25 * i, 2) for i in range(21)] + [21.0]   # 8.00 .. 13.00, wide diagnostic
WIDTHS_ARIAL = [round(10.0 + 0.25 * i, 2) for i in range(25)]  # 10.00 .. 16.00
WIDTHS_LITERAL = [round(10.0 + 0.5 * i, 2) for i in range(11)]  # 10.00 .. 15.00

# name, face, size, italic, align, scale, kind, widths
CASES = [
    ('R_aptos14_upright', 'Aptos Narrow', 14, False, 'right',  100, 'overflow', WIDTHS_APTOS),
    ('L_aptos14_upright', 'Aptos Narrow', 14, False, 'left',   100, 'overflow', WIDTHS_APTOS),
    ('C_aptos14_upright', 'Aptos Narrow', 14, False, 'center', 100, 'overflow', WIDTHS_APTOS),
    ('R_aptos14_italic',  'Aptos Narrow', 14, True,  'right',  100, 'overflow', WIDTHS_APTOS),
    ('L_aptos14_italic',  'Aptos Narrow', 14, True,  'left',   100, 'overflow', WIDTHS_APTOS),
    ('C_aptos14_italic',  'Aptos Narrow', 14, True,  'center', 100, 'overflow', WIDTHS_APTOS),
    ('R_literal_aptos14', 'Aptos Narrow', 14, False, 'right',  100, 'literal',  WIDTHS_LITERAL),
    ('L_literal_aptos14', 'Aptos Narrow', 14, False, 'left',   100, 'literal',  WIDTHS_LITERAL),
    ('C_literal_aptos14', 'Aptos Narrow', 14, False, 'center', 100, 'literal',  WIDTHS_LITERAL),
    ('R_literal_aptos14_i','Aptos Narrow',14, True,  'right',  100, 'literal',  WIDTHS_LITERAL),
    # held out: slopes are font- and scale-independent; counts/advances are not predicted from constants
    ('R_calibri11',       'Calibri',      11, False, 'right',  100, 'overflow', WIDTHS_APTOS),
    ('R_calibri11_italic','Calibri',      11, True,  'right',  100, 'overflow', WIDTHS_APTOS),
    ('R_arial12',         'Arial',        12, False, 'right',  100, 'overflow', WIDTHS_ARIAL),
    ('R_aptos14_75',      'Aptos Narrow', 14, False, 'right',   75, 'overflow', WIDTHS_APTOS),
    ('R_aptos14_italic_75','Aptos Narrow',14, True,  'right',   75, 'overflow', WIDTHS_APTOS),
]


def attr(text):
    return escape(text, {'"': '&quot;'})


def fill_color(i):
    """A distinct, decidedly non-white/non-black solid fill for each test column index."""
    return '%02X%02X%02X' % (0xE0 + (i % 16), 0x20 + ((i * 37) % 0xC0), 0x30 + ((i * 71) % 0xA0))


def rgb_text(argb):
    return '%d,%d,%d' % (int(argb[0:2], 16), int(argb[2:4], 16), int(argb[4:6], 16))


def build(out, meta_path):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, *_ in CASES:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')
    # fill 0 none, fill 1 gray125, fills 2.. one per column index
    max_cols = max(len(c[-1]) for c in CASES)
    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    for i in range(max_cols):
        rgb = fill_color(i)
        fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/><bgColor indexed="64"/></patternFill></fill>')

    # cellXfs: 0 default, 1 label; then one per (case, column, row)
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>']
    style_of = {}
    for ci, (_, _, _, italic, align, _, kind, widths) in enumerate(CASES):
        font_id = ci + 1
        for col, _w in enumerate(widths):
            for row, fmt in ((1, 1 if kind == 'overflow' else 49), (2, 200 if kind == 'overflow' else 49)):
                style_of[(ci, col, row)] = len(xfs)
                xfs.append(f'<xf numFmtId="{fmt}" fontId="{font_id}" fillId="{2 + col}" borderId="0" xfId="0" '
                           f'applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                           f'<alignment horizontal="{align}"/></xf>')
    numfmts = '<numFmt numFmtId="200" formatCode="0.00"/>'
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><numFmts count="1">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(fills)}">' + ''.join(fills) + '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    sheet_files = []
    meta = ['sheet\tcolumn\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\trow1\trow2']
    for ci, (name, face, size, italic, align, scale, kind, widths) in enumerate(CASES):
        cols = ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>' for j, w in enumerate(widths))
        cells = ['<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>%s</t></is></c>' % escape(name)]
        for j, w in enumerate(widths):
            c = chr(ord('B') + j) if j < 25 else None
            ref = f'{c}1' if c else None
            if c is None:
                raise SystemExit('too many columns for this simple A1 writer')
            s = style_of[(ci, j, 1)]
            if kind == 'overflow':
                cells.append(f'<c r="{ref}" s="{s}"><v>1123456789012</v></c>')
            else:
                cells.append(f'<c r="{ref}" s="{s}" t="inlineStr"><is><t>#######</t></is></c>')
        cells.append('</row><row r="2">')
        for j, w in enumerate(widths):
            c = chr(ord('B') + j)
            s = style_of[(ci, j, 2)]
            if kind == 'overflow':
                cells.append(f'<c r="{c}2" s="{s}"><v>-98.41</v></c>')
                row2 = '-98.41'
            else:
                cells.append(f'<c r="{c}2" s="{s}" t="inlineStr"><is><t>########</t></is></c>')
                row2 = '########'
            meta.append('\t'.join([name, str(2 + j), f'{w:g}', align, face, str(size),
                                   '1' if italic else '0', str(scale), kind,
                                   '1123456789012' if kind == 'overflow' else '#######', row2]))
        cells.append('</row>')
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if ci == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="22" customWidth="1"/>{cols}</cols>'
            f'<sheetData>{"".join(cells)}</sheetData>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheet_files))
            + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                      for i in range(len(sheet_files)))
            + f'<Relationship Id="rId{len(sheet_files) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheet_files):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    with open(meta_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(meta) + '\n')


if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1].rsplit('.', 1)[0] + '.meta.tsv')
