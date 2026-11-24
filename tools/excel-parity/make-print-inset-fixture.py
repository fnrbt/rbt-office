#!/usr/bin/env python3
"""Plain-number print insets: number advance and left/right text insets versus cell font, size and scale.

No fill, no border. The workbook Normal font is fixed (Calibri 11 unless `--normal-face/--normal-size`), column widths are fixed, and each
sheet is one (print scale, horizontal alignment). Every row is one (cell font, cell size, kind):

  pos   1234.5 under `0.00`  -> "1234.50"
  neg  -1234.5 under `0.00`  -> "-1234.50"
  zero       0 under `0.00`  -> "0.00"
  text  the string "1234.50" (General), the same glyphs as `pos` drawn as text

Cell geometry comes from per-column solid fills; the measurement reads the text origin and advance end
from the PDF operators, so left inset = x0 - cell_left (left aligned) and right inset =
cell_right - x1 (right aligned).

The full grid is FACES x SIZES x SCALES. `--set train` writes the training combinations, `--set holdout`
the rest, so predictions can be locked on the training export before the holdout is exported.

    python3 tools/excel-parity/make-print-inset-fixture.py --set train --out BOOK.xlsx --cells CELLS.tsv

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

FACES = ['Aptos Narrow', 'Arial', 'Calibri', 'Georgia', 'Tahoma', 'Verdana']
SIZES = [9, 10, 11, 12, 14]
SCALES = [60, 75, 100]
TRAIN_FACES = {'Aptos Narrow', 'Arial', 'Calibri', 'Georgia'}
TRAIN_SIZES = {9, 11, 14}
TRAIN_SCALES = {60, 100}
KINDS = [('pos', 'n', '1234.5'), ('neg', 'n', '-1234.5'), ('zero', 'n', '0'), ('text', 's', '1234.50')]
ALIGNS = ['left', 'center', 'right']
WIDTHS = ['14', '18']
PALETTE = ['DDEBF7', 'FFF2CC']


def in_train(face, size, scale):
    return face in TRAIN_FACES and size in TRAIN_SIZES and scale in TRAIN_SCALES


def attr(text):
    return escape(text, {'"': '&quot;'})


def build(which, out, cells_path, normal_face='Calibri', normal_size=11):
    combos = [(f, s, sc) for sc in SCALES for f in FACES for s in SIZES
              if in_train(f, s, sc) == (which == 'train')]
    scales = sorted({sc for _, _, sc in combos})
    fonts = [f'<sz val="{normal_size:g}"/><name val="{attr(normal_face)}"/><family val="2"/>']
    font_id = {}
    for f, s, _ in combos:
        if (f, s) not in font_id:
            font_id[(f, s)] = len(fonts)
            fonts.append(f'<sz val="{s}"/><name val="{attr(f)}"/><family val="2"/>')
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             + ''.join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{c}"/>'
                       '<bgColor indexed="64"/></patternFill></fill>' for c in PALETTE))
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for (f, s), fid in font_id.items():
        for kind, typ, _ in KINDS:
            for align in ALIGNS:
                for wi in range(len(WIDTHS)):
                    style[(f, s, kind, align, wi)] = len(xfs)
                    num = '2' if typ == 'n' else '0'   # built-in 2 is `0.00`
                    xfs.append(
                        f'<xf numFmtId="{num}" fontId="{fid}" fillId="{2 + wi}" borderId="0" xfId="0" '
                        'applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                        f'<alignment horizontal="{align}"/></xf>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{x}</font>' for x in fonts) + '</fonts>'
        f'<fills count="{2 + len(PALETTE)}">{fills}</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    sheets = [(f'P{sc}_{align}', sc, align) for sc in scales for align in ALIGNS]
    sheet_files = []
    manifest = []
    for si, (sname, scale, align) in enumerate(sheets):
        rows = []
        head = '<row r="1"><c r="A1" t="inlineStr"><is><t>width</t></is></c>'
        head += ''.join(f'<c r="{chr(66 + j)}1" t="inlineStr"><is><t>{w}</t></is></c>'
                        for j, w in enumerate(WIDTHS))
        rows.append(head + '</row>')
        r = 2
        for f, s, sc in combos:
            if sc != scale:
                continue
            for kind, typ, value in KINDS:
                label = f'r{r:03d}'
                cells = ''
                for wi in range(len(WIDTHS)):
                    ref = f'{chr(66 + wi)}{r}'
                    sid = style[(f, s, kind, align, wi)]
                    if typ == 'n':
                        cells += f'<c r="{ref}" s="{sid}"><v>{value}</v></c>'
                    else:
                        cells += f'<c r="{ref}" s="{sid}" t="inlineStr"><is><t>{value}</t></is></c>'
                rows.append(f'<row r="{r}" ht="21" customHeight="1">'
                            f'<c r="A{r}" t="inlineStr"><is><t>{label}</t></is></c>{cells}</row>')
                manifest.append((sname, si + 1, r, label, scale, align, f, s, kind, value))
                r += 1
        cols = '<col min="1" max="1" width="8" customWidth="1"/>'
        cols += ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                        for j, w in enumerate(WIDTHS))
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if si == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
            '<pageMargins left="0.25" right="0.25" top="0.3" bottom="0.3" header="0.15" footer="0.15"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
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
        f.write('sheet\tindex\trow\tlabel\tscale\talign\tface\tsize\tkind\tvalue\twidths\n')
        for sname, index, row, label, scale, align, face, size, kind, value in manifest:
            f.write(f'{sname}\t{index}\t{row}\t{label}\t{scale}\t{align}\t{face}\t{size}\t{kind}\t{value}\t'
                    f'{",".join(WIDTHS)}\n')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--set', choices=['train', 'holdout'], required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--cells', required=True)
    ap.add_argument('--normal-face', default='Calibri')
    ap.add_argument('--normal-size', type=float, default=11)
    args = ap.parse_args()
    build(args.set, args.out, args.cells, args.normal_face, args.normal_size)


if __name__ == '__main__':
    main()
