#!/usr/bin/env python3
"""Discriminating border/fill fixture: trailing `*x` fill boundary versus border weight.

Identical number and trailing-fill/no-fill twin rows under no / thin / medium / thick cell borders,
with the same column widths in every sheet of a workbook. Used to separate three competing models
of the fill repeat boundary:

  half-border   the fill repeats to the border's inner edge, so the usable width gains half the
                physical border width (center -> inner edge);
  fixed-2-dot   the usable width gains a fixed 0.24 pt (2 device dots at 600 dpi) regardless of
                border weight;
  none          the usable width is the no-fill text content width (no correction).

Sheets: B<none|thin|medium|thick>_<left|centre|right>, 12 columns of 0.25-char width steps that
bracket count transitions for all three values. Rows: `0.00` control, `0.00* `, `0.00*.`,
`0.00*-;0.00`, positive/negative/zero. Borders are square boxes in gray on every test cell; every
test cell also has its own solid fill so the no-border sheet still measures cell geometry from the
fill rectangle edges.

    python3 tools/excel-parity/make-print-border-fill-fixture.py --face Arial --size 11 --scale 60 \
        --out BOOK.xlsx --cells CELLS.tsv

Deterministic: bytes depend only on the arguments.
"""
import argparse
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

FORMATS = [
    ('N1_plain_0', '0.00'),
    ('S2_space_after', '0.00* '),
    ('D2_dot_after', '0.00*.'),
    ('F8_suffix_empty', '0.00*-;0.00'),
]
VALUES = [('pos', '1234.5'), ('neg', '-1234.5'), ('zero', '0')]
ALIGNS = ['left', 'center', 'right']
BORDERS = ['none', 'thin', 'medium', 'thick']
WIDTHS = ['8.5', '8.75', '9', '9.25', '9.5', '9.75', '10', '10.25', '10.5', '10.75', '11', '11.25']
PALETTE = ['DDEBF7', 'FFF2CC', 'E2EFDA', 'FCE4D6', 'E4DFEC', 'D9E1F2', 'EDEDED', 'FFE699',
           'D9F2E6', 'F2E6D9', 'E6D9F2', 'F2D9E6']


def attr(text):
    return escape(text, {'"': '&quot;'})


def border_xml(style):
    if style == 'none':
        return ('<border><left/><right/><top/><bottom/><diagonal/></border>')
    edge = f'style="{style}"><color rgb="FF808080"/>'
    return ('<border>'
            + ''.join(f'<{side} {edge}</{side}>' for side in ('left', 'right', 'top', 'bottom'))
            + '<diagonal/></border>')


def build(out, cells_path, face, size, scale):
    sheets = [(f'B{border}_{align}', border, align) for border in BORDERS for align in ALIGNS]
    fonts = [
        '<sz val="11"/><name val="Calibri"/><family val="2"/>',
        f'<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>',
    ]
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{c}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>' for c in PALETTE))
    borders = '<borders count="4">' + ''.join(border_xml(b) for b in BORDERS) + '</borders>'
    border_id = {b: i for i, b in enumerate(BORDERS)}
    code_id = {}
    numfmts = ''
    for _, code in FORMATS:
        code_id[code] = 200 + len(code_id)
        numfmts += f'<numFmt numFmtId="{code_id[code]}" formatCode="{attr(code)}"/>'

    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for si, (sname, border, align) in enumerate(sheets):
        for fi, (_, code) in enumerate(FORMATS):
            for vi in range(len(VALUES)):
                for wi in range(len(WIDTHS)):
                    style[(si, fi, vi, wi)] = len(xfs)
                    xfs.append(
                        f'<xf numFmtId="{code_id[code]}" fontId="1" fillId="{2 + wi}" '
                        f'borderId="{border_id[border]}" xfId="0" applyFont="1" applyFill="1" '
                        f'applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                        f'<alignment horizontal="{align}"/></xf>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><numFmts count="{len(code_id)}">{numfmts}</numFmts>'
        f'<fonts count="2">' + ''.join(f'<font>{x}</font>' for x in fonts) + '</fonts>'
        f'<fills count="{2 + len(PALETTE)}">{fills}</fills>{borders}'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    sheet_files = []
    manifest = []
    for si, (sname, border, align) in enumerate(sheets):
        rows = []
        head = '<row r="1"><c r="A1" t="inlineStr"><is><t>width</t></is></c>'
        head += ''.join(f'<c r="{chr(66 + j)}1" t="inlineStr"><is><t>{w}</t></is></c>'
                        for j, w in enumerate(WIDTHS))
        rows.append(head + '</row>')
        r = 2
        for fi, (fname, code) in enumerate(FORMATS):
            for vi, (vname, value) in enumerate(VALUES):
                cells = ''
                for wi in range(len(WIDTHS)):
                    cells += (f'<c r="{chr(66 + wi)}{r}" s="{style[(si, fi, vi, wi)]}">'
                              f'<v>{value}</v></c>')
                rows.append(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{sname} {fname} {vname}</t></is></c>{cells}</row>')
                manifest.append((sname, si + 1, r, fname, code, value, WIDTHS, align))
                r += 1
        cols = '<col min="1" max="1" width="20" customWidth="1"/>'
        cols += ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                        for j, w in enumerate(WIDTHS))
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if si == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
            '<pageMargins left="0.25" right="0.25" top="0.3" bottom="0.3" header="0.15" footer="0.15"/>'
            f'<pageSetup paperSize="1" orientation="landscape" scale="{scale}"/></worksheet>')
        sheet_files.append((sname, xml))
    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + (
            f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (n, _) in enumerate(sheet_files)) + '</sheets></workbook>'),
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
    with open(cells_path, 'w', encoding='utf-8') as f:
        f.write('sheet\tindex\trow\tformat_name\tformat\tvalue\twidths\talign\n')
        for sname, index, row, fname, code, value, widths, align in manifest:
            f.write(f'{sname}\t{index}\t{row}\t{fname}\t{code}\t{value}\t{",".join(widths)}\t{align}\n')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--face', required=True)
    ap.add_argument('--size', type=float, required=True)
    ap.add_argument('--scale', type=int, required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--cells', required=True)
    args = ap.parse_args()
    build(args.out, args.cells, args.face, args.size, args.scale)


if __name__ == '__main__':
    main()
