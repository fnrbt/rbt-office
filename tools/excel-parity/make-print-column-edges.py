#!/usr/bin/env python3
"""Builds an independent vector-geometry fixture for printed spreadsheet column edges.

Why: the printed full/hash cutoff around Georgia 15 upright 90 % columns 24-28
(widths 16.2188-16.3438 characters) misses by 0.04-0.15 pt when the layout width
is used, and a previously measured PDF *fill-rectangle* width is 0.07-0.20 pt
wider than the layout width.  A fill rectangle can be expanded for coverage, so
it is not by itself the cell boundary.  This workbook makes the actual boundary
visible and measurable in the PDF content stream:

  * each test column has a unique solid fill (the fill rectangle), and
  * every test cell has a thin black border on all four sides (the grid line).

The fill path and the border path are different vector objects in the exported
PDF, so subtracting them separates cell width from fill/stroke expansion.

Sheets:
  E90_edge    Georgia 15 upright, 90 % print scale, test columns start at B.
              40 repeated equal-width columns (16.2812 chars) followed by 41
              varied widths 15.75 .. 17.00 by 1/32 char (the five failure
              columns are interior), so cumulative origins, page breaks and
              widths can all be compared in one pass.
  H85_edge    HOLD-OUT: same layout at 85 % print scale, margins 0.25 in.
  H90_shift   HOLD-OUT: 90 %, but test columns start at E; A-D have widths
              8/22/13/9 chars, and margins are 0.7 in left / 0.25 in right.
              This changes the page x origin without changing the widths.

Rows: 1 = 1234567 with format #,##0.00 (may print the full number or hashes),
2 = -98.41 with format 0.00 (never overflows), 3 = a short index label.
Column A carries the sheet label.  The workbook is written directly as OOXML
and must be exported by excel-print-pages.sh (direct Excel 2019 PDF before any
activation/preview).  META.tsv records every fill colour, width, margin and the
pre-registered layout width for offline comparison.

    python3 tools/excel-parity/make-print-column-edges.py OUT.xlsx META.tsv
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

FACE, SIZE, ITALIC = 'Georgia', 15, False
REPEAT_W = 16.2812
REPEAT_N = 40
VARY = [round(15.75 + i / 32.0, 4) for i in range(41)]     # 15.75 .. 17.00
PREFIX = {'E90_edge': [22.0], 'H85_edge': [22.0],
          'H90_shift': [8.0, 22.0, 13.0, 9.0]}
# name, scale, left margin in, right margin in, start column (1-based)
SHEETS = [
    ('E90_edge', 90, 0.4, 0.4, 2),
    ('H85_edge', 85, 0.25, 0.25, 2),
    ('H90_shift', 90, 0.7, 0.25, 5),
]


def fill_color(i):
    """Unique-ish solid colour per test column; deliberately far from black/white."""
    r = 40 + (i * 67) % 190
    g = 40 + (i * 113) % 190
    b = 40 + (i * 157) % 190
    return r, g, b


def col_letter(c):
    s = ''
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def layout_width(width_chars, scale):
    """The product's stored-width -> printed-points conversion at a print scale."""
    screen_px = int(((256.0 * width_chars + 128 // 7) / 256.0) * 7.0)
    return int(screen_px * 47.0 / 7.0 + 0.5) * 72.0 / 600.0 * scale / 100.0


def attr(text):
    return escape(text, {'"': '&quot;'})


def build(out, meta_path):
    sheets = []
    meta = ['sheet\tcolumn\tblock\tblock_index\twidth_chars\tscale\tmargin_left_in\tmargin_right_in\t'
            'face\tsize_pt\titalic\trow1_format\trow1_value\trow2_format\trow2_value\tfill_rgb\tborder']
    all_test = []
    for name, scale, ml, mr, start_col in SHEETS:
        cols = []
        for j in range(REPEAT_N):
            cols.append(('repeat', j, REPEAT_W))
        for j, w in enumerate(VARY):
            cols.append(('vary', j, w))
        all_test.append((name, scale, ml, mr, start_col, cols))

    max_fill = max(len(t[5]) for t in all_test)
    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    for i in range(max_fill):
        r, g, b = fill_color(i)
        fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{r:02X}{g:02X}{b:02X}"/>'
                     f'<bgColor indexed="64"/></patternFill></fill>')

    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>',          # 0 default
             f'<sz val="{SIZE}"/><name val="{attr(FACE)}"/><family val="2"/>',  # 1 test numbers
             '<sz val="9"/><name val="Calibri"/><family val="2"/>']            # 2 index labels
    borders = ('<borders count="2">'
               '<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border>'
               '<left style="thin"><color rgb="FF000000"/></left>'
               '<right style="thin"><color rgb="FF000000"/></right>'
               '<top style="thin"><color rgb="FF000000"/></top>'
               '<bottom style="thin"><color rgb="FF000000"/></bottom>'
               '<diagonal/></border></borders>')

    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
           '<alignment horizontal="left"/></xf>']
    style_of = {}
    for sheet_i, (name, scale, ml, mr, start_col, cols) in enumerate(all_test):
        for j, (block, bidx, w) in enumerate(cols):
            fill_id = 2 + j
            for row, font, fmt, align in (
                    (1, 1, 167, 'right'), (2, 1, 2, 'right'), (3, 2, 0, 'left')):
                style_of[(sheet_i, j, row)] = len(xfs)
                xfs.append(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill_id}" borderId="1" '
                           f'xfId="0" applyFont="1" applyFill="1" applyBorder="1" '
                           f'applyNumberFormat="1" applyAlignment="1">'
                           f'<alignment horizontal="{align}"/></xf>')

    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        '<numFmts count="1"><numFmt numFmtId="167" formatCode="#,##0.00"/></numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(fills)}">' + ''.join(fills) + '</fills>'
        f'{borders}'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')

    sheet_files = []
    for sheet_i, (name, scale, ml, mr, start_col, cols) in enumerate(all_test):
        first = start_col
        last = start_col + len(cols) - 1
        prefix = PREFIX[name]
        assert len(prefix) == start_col - 1, (name, len(prefix), start_col)
        col_xml = ''.join(
            f'<col min="{1 + j}" max="{1 + j}" width="{w}" customWidth="1"/>'
            for j, w in enumerate(prefix))
        col_xml += ''.join(
            f'<col min="{start_col + j}" max="{start_col + j}" width="{w}" customWidth="1"/>'
            for j, (_, _, w) in enumerate(cols))

        rows = []
        rows.append(f'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>{escape(name)}</t></is></c>')
        for j, _ in enumerate(cols):
            ref = f'{col_letter(start_col + j)}1'
            rows.append(f'<c r="{ref}" s="{style_of[(sheet_i, j, 1)]}"><v>1234567</v></c>')
        rows.append('</row>')
        rows.append('<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>control</t></is></c>')
        for j, _ in enumerate(cols):
            ref = f'{col_letter(start_col + j)}2'
            rows.append(f'<c r="{ref}" s="{style_of[(sheet_i, j, 2)]}"><v>-98.41</v></c>')
        rows.append('</row>')
        rows.append('<row r="3"><c r="A3" s="1" t="inlineStr"><is><t>index</t></is></c>')
        for j, (block, bidx, w) in enumerate(cols):
            ref = f'{col_letter(start_col + j)}3'
            label = f'{block[:1].upper()}{bidx:02d}'
            rows.append(f'<c r="{ref}" s="{style_of[(sheet_i, j, 3)]}" t="inlineStr"><is><t>{label}</t></is></c>')
        rows.append('</row>')

        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if sheet_i == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols>{col_xml}</cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            f'<pageMargins left="{ml}" right="{mr}" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

        for j, (block, bidx, w) in enumerate(cols):
            r, g, b = fill_color(j)
            meta.append('\t'.join([
                name, str(start_col + j), block, str(bidx), f'{w:g}', str(scale),
                f'{ml:g}', f'{mr:g}', FACE, str(SIZE), '0', '#,##0.00', '1234567', '0.00', '-98.41',
                f'{r:02X}{g:02X}{b:02X}', 'thin']))

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
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (name, _) in enumerate(sheet_files))
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
    print(f'wrote {out}: {len(sheet_files)} sheets, {sum(len(t[5]) for t in all_test)} test columns')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2])
