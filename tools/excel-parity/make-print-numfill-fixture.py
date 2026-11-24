#!/usr/bin/env python3
"""Write the numeric-fill print oracle workbooks (Excel 2019 direct PDF layout).

The workbook tests what Excel does with an active `*x` fill token in a number format, compared with
matched no-fill formats: the segment before the fill, the repeated fill character, and the segment
after it; positive/negative/zero sections; left/centre/right cell alignment; tight column-width steps
across repeat transitions; quoted/escaped `*` and inactive sections. Each test cell has its own solid
fill (a distinct colour per width column) and thin borders, so the measurement can read the physical
cell edge and the PDF text operators independently of page geometry.

Sets:
  train           Aptos Narrow 14 at 100%: short formats (widths 9.25..11.0) and wide formats
                  (13.0..14.75), each on left/centre/right sheets.
  holdout-georgia Georgia 12 at 75%: the same faces reduced to the formats a blind rule must cover.
  holdout-tahoma  Tahoma 10 at 100%: the same reduced set at a second face/scale.

    python3 tools/excel-parity/make-print-numfill-fixture.py --set train --out BOOK.xlsx --cells CELLS.tsv

Deterministic: the bytes depend only on the arguments.
"""
import argparse, zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

# (name, format code). A leading `S_`/`D_` is a short-group format; `W_` a wide-group format.
SHORT_FORMATS = [
    ('N0_plain', '#,##0.00'),
    ('N1_plain_0', '0.00'),
    ('S1_space_before', '* #,##0.00'),
    ('D1_dot_before', '*.#,##0.00'),
    ('S2_space_after', '0.00* '),
    ('D2_dot_after', '0.00*.'),
    ('Q1_quoted_star', '"*"#,##0.00'),
    ('E1_escaped_star', r'\*#,##0.00'),
    ('I1_fill_neg_only', '0.00;[Red]* -0.00;0.00'),
    ('I2_fill_pos_only', '0.00*.;0.00'),
]
WIDE_FORMATS = [
    ('W_C1_currency_space', '$* #,##0.00'),
    ('W_C2_currency_dot', '$*.#,##0.00'),
    ('W_A1_accounting', '_($* #,##0.00_);_($* \\(#,##0.00\\);_($* "-"??_);_(@_)'),
    ('W_A2_accounting_dot', '_($*.#,##0.00_);_($* \\(#,##0.00\\);_($* "-"??_);_(@_)'),
    ('W_I3_currency_neg_only', '$0.00;[Red]$* -0.00;0.00'),
]
HOLDOUT_SHORT = ['N0_plain', 'N1_plain_0', 'S1_space_before', 'D1_dot_before', 'S2_space_after',
                 'Q1_quoted_star', 'I1_fill_neg_only']
HOLDOUT_WIDE = ['W_C1_currency_space', 'W_A1_accounting']
SHORT_WIDTHS = ['9.25', '9.5', '9.75', '10', '10.25', '10.5', '10.75', '11']
WIDE_WIDTHS = ['13', '13.25', '13.5', '13.75', '14', '14.25', '14.5']
# Trailing-fill sweep: suffix-empty `*x` where the fill token ends the active section. Widths run
# from where `0.00` itself stops fitting up through several fill-count transitions; each band is a
# separate sheet because a landscape page cannot carry the whole 0.25-char sweep with its header.
TRAIL_FORMATS = [
    ('N1_plain_0', '0.00'),
    ('S2_space_after', '0.00* '),
    ('D2_dot_after', '0.00*.'),
    ('F8_suffix_empty', '0.00*-;0.00'),
]
TRAIL_WIDTH_BANDS = [
    ['5.5', '5.75', '6', '6.25', '6.5', '6.75', '7'],
    ['7.25', '7.5', '7.75', '8', '8.25', '8.5', '8.75'],
    ['9', '9.25', '9.5', '9.75', '10', '10.25', '10.5'],
    ['10.75', '11', '11.25', '11.5', '11.75', '12'],
]
VALUES = [('pos', '1234.5'), ('neg', '-1234.5'), ('zero', '0')]
ALIGNS = ['left', 'center', 'right']
PALETTE = ['DDEBF7', 'FFF2CC', 'E2EFDA', 'FCE4D6', 'E4DFEC', 'D9E1F2', 'EDEDED', 'FFE699']


def attr(text):
    return escape(text, {'"': '&quot;'})


def build_set(name, face, size, scale):
    if name == 'trail-train' or name == 'trail-holdout':
        return [
            (f'T{band + 1}_{align}', TRAIL_FORMATS, widths, align)
            for band, widths in enumerate(TRAIL_WIDTH_BANDS)
            for align in ALIGNS
        ]
    if name == 'train':
        short = SHORT_FORMATS
        wide = WIDE_FORMATS
        sw, ww = SHORT_WIDTHS, WIDE_WIDTHS
    else:
        short_map = dict(SHORT_FORMATS)
        wide_map = dict(WIDE_FORMATS)
        short = [(n, short_map[n]) for n in HOLDOUT_SHORT]
        wide = [(n, wide_map[n]) for n in HOLDOUT_WIDE]
        sw, ww = SHORT_WIDTHS, WIDE_WIDTHS
    return [
        (f'{prefix}_{align}', formats, widths, align)
        for prefix, formats, widths in (('S', short, sw), ('W', wide, ww))
        for align in ALIGNS
    ]


def build(out, cells_path, set_name, face, size, scale):
    sheets = build_set(set_name, face, size, scale)
    fonts = [
        '<sz val="11"/><name val="Calibri"/><family val="2"/>',
        f'<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>',
    ]
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{c}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>' for c in PALETTE))
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border><left style="thin"><color rgb="FF808080"/></left>'
               '<right style="thin"><color rgb="FF808080"/></right>'
               '<top style="thin"><color rgb="FF808080"/></top>'
               '<bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for si, (sname, formats, widths, align) in enumerate(sheets):
        for fi, (fname, code) in enumerate(formats):
            for vi, (vname, value) in enumerate(VALUES):
                for wi in range(len(widths)):
                    style[(si, fi, vi, wi)] = len(xfs)
                    xfs.append(
                        f'<xf numFmtId="{200 + fi}" fontId="1" fillId="{2 + wi}" borderId="1" xfId="0" '
                        f'applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                        f'<alignment horizontal="{align}"/></xf>')
    numfmts = ''
    seen = set()
    for _, formats, _, _ in sheets:
        for fname, code in formats:
            if code not in seen:
                seen.add(code)
                numfmts += f'<numFmt numFmtId="{200 + len(seen) - 1}" formatCode="{attr(code)}"/>'
    # The style index above assumes one numFmt per format position in each group; groups repeat the
    # same codes, so map code -> id and rebuild xfs deterministically.
    code_id = {}
    numfmts = ''
    for _, formats, _, _ in sheets:
        for _, code in formats:
            if code not in code_id:
                code_id[code] = 200 + len(code_id)
                numfmts += f'<numFmt numFmtId="{code_id[code]}" formatCode="{attr(code)}"/>'
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for si, (sname, formats, widths, align) in enumerate(sheets):
        for fi, (fname, code) in enumerate(formats):
            for vi in range(len(VALUES)):
                for wi in range(len(widths)):
                    style[(si, fi, vi, wi)] = len(xfs)
                    xfs.append(
                        f'<xf numFmtId="{code_id[code]}" fontId="1" fillId="{2 + wi}" borderId="1" xfId="0" '
                        f'applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                        f'<alignment horizontal="{align}"/></xf>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><numFmts count="{len(code_id)}">{numfmts}</numFmts>'
        f'<fonts count="2">' + ''.join(f'<font>{x}</font>' for x in fonts) + '</fonts>'
        f'<fills count="{2 + len(PALETTE)}">{fills}</fills><borders count="2">{borders}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    sheet_files = []
    manifest = []
    for si, (sname, formats, widths, align) in enumerate(sheets):
        rows = []
        head = '<row r="1"><c r="A1" t="inlineStr"><is><t>width</t></is></c>'
        head += ''.join(f'<c r="{chr(66 + j)}1" t="inlineStr"><is><t>{w}</t></is></c>'
                        for j, w in enumerate(widths))
        rows.append(head + '</row>')
        r = 2
        for fi, (fname, code) in enumerate(formats):
            for vi, (vname, value) in enumerate(VALUES):
                cells = ''
                for wi in range(len(widths)):
                    cells += (f'<c r="{chr(66 + wi)}{r}" s="{style[(si, fi, vi, wi)]}">'
                              f'<v>{value}</v></c>')
                rows.append(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{sname} {fname} {vname}</t></is></c>{cells}</row>')
                manifest.append((sname, si + 1, r, fname, code, value, widths, align))
                r += 1
        cols = '<col min="1" max="1" width="20" customWidth="1"/>'
        cols += ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                        for j, w in enumerate(widths))
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
    ap.add_argument('--set', required=True,
                    choices=['train', 'holdout-georgia', 'holdout-tahoma', 'trail-train', 'trail-holdout'])
    ap.add_argument('--out', required=True)
    ap.add_argument('--cells', required=True)
    ap.add_argument('--face')
    ap.add_argument('--size', type=float)
    ap.add_argument('--scale', type=int)
    args = ap.parse_args()
    defaults = {
        'train': ('Aptos Narrow', 14.0, 100),
        'holdout-georgia': ('Georgia', 12.0, 75),
        'holdout-tahoma': ('Tahoma', 10.0, 100),
        'trail-train': ('Aptos Narrow', 14.0, 100),
        'trail-holdout': ('Arial', 11.0, 60),
    }[args.set]
    face = args.face or defaults[0]
    size = args.size or defaults[1]
    scale = args.scale or defaults[2]
    build(args.out, args.cells, args.set, face, size, scale)


if __name__ == '__main__':
    main()
