#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) for where Excel 2019 sets a line of cell text in its cell on
screen. Each workbook has one Normal font and sheets at zooms; every case is three cells - the case between two
cells filled black (markers for the cell's edges) - with a low empty row between rows of cases and up to three
cases a row. Shot in Excel (`shoot-sheets.ps1 -KeepZoom`) and in the browser grid (`shoot-browser-sheets.mjs`), the
ink of each case compared (`measure-text-inset.mjs`: the shift that lays ours over Excel's).
docs/excel-parity-tracker.md, "Text insets".

    python3 tools/excel-parity/make-text-inset-fixture.py OUTDIR [NAME ...]

Writes OUTDIR/inset-<name>.xlsx for each workbook named (all when none are), and prints a line per sheet
(path, sheet) and, under it, the sheet's cases in the order the measure lists them. Deterministic.

- "base" workbooks (Normal fonts Calibri 11 ... Arial 28; zooms): alignments of the Normal font's text - general,
  left, right, center, indents 1 and 3, fill, distributed, justify, numbers - and texts in other fonts.
- "sweep" workbooks (Normal Calibri 11): a font at many sizes, each a general text, a right-aligned text and a
  General number.
"""
import os
import sys
import zipfile

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
TEXT = 'HHHH'


def base_cases(normal):
    n = normal
    return [
        ('general', '', n, TEXT, False), ('left', 'horizontal="left"', n, TEXT, False),
        ('right', 'horizontal="right"', n, TEXT, False), ('center', 'horizontal="center"', n, TEXT, False),
        ('left-i1', 'horizontal="left" indent="1"', n, TEXT, False), ('right-i1', 'horizontal="right" indent="1"', n, TEXT, False),
        ('left-i3', 'horizontal="left" indent="3"', n, TEXT, False), ('fill', 'horizontal="fill"', n, 'H', False),
        ('distributed', 'horizontal="distributed"', n, 'H H H', False),
        ('dist-i1', 'horizontal="distributed" indent="1"', n, 'H H H', False),
        ('justify', 'horizontal="justify"', n, TEXT, False), ('num-general', '', n, '1000', True),
        ('num-left', 'horizontal="left"', n, '1000', True), ('num-right', 'horizontal="right"', n, '1000', True),
        ('num-center', 'horizontal="center"', n, '1000', True), ('fill-2', 'horizontal="fill"', n, 'Hx', False),
        ('f-calibri11-g', '', ('Calibri', 11, False), TEXT, False), ('f-calibri11-r', 'horizontal="right"', ('Calibri', 11, False), TEXT, False),
        ('f-arial20-g', '', ('Arial', 20, False), TEXT, False), ('f-arial20-r', 'horizontal="right"', ('Arial', 20, False), TEXT, False),
        ('f-arial20-n', '', ('Arial', 20, False), '1000', True), ('f-bold-g', '', ('Calibri', 11, True), TEXT, False),
        ('wrap-l', 'horizontal="left" wrapText="1"', n, TEXT, False), ('wrap-r', 'horizontal="right" wrapText="1"', n, TEXT, False),
    ]


def sweep_cases(face, sizes, bold=False):
    out = []
    for s in sizes:
        f = (face, s, bold)
        out += [(f'{s}-g', '', f, TEXT, False), (f'{s}-r', 'horizontal="right"', f, TEXT, False), (f'{s}-n', '', f, '1000', True)]
    return out


CALIBRI11 = ('Calibri', 11, False)
# name: (Normal font, [(sheet, zoom, cases, cases a row)])
BOOKS = {
    'calibri11': (CALIBRI11, [(f'Z{z}', z, base_cases(CALIBRI11), 2) for z in (100, 75, 90, 150, 200)]),
    'arial10': (('Arial', 10, False), [(f'Z{z}', z, base_cases(('Arial', 10, False)), 2) for z in (100, 150)]),
    'calibri20': (('Calibri', 20, False), [(f'Z{z}', z, base_cases(('Calibri', 20, False)), 2) for z in (100, 75)]),
    'cambria14': (('Cambria', 14, False), [('Z100', 100, base_cases(('Cambria', 14, False)), 2)]),
    'verdana14': (('Verdana', 14, False), [('Z100', 100, base_cases(('Verdana', 14, False)), 2)]),
    'times8': (('Times New Roman', 8, False), [('Z100', 100, base_cases(('Times New Roman', 8, False)), 2)]),
    'segoe9': (('Segoe UI', 9, False), [('Z100', 100, base_cases(('Segoe UI', 9, False)), 2)]),
    'sweep': (CALIBRI11, [
        ('Calibri', 100, sweep_cases('Calibri', (8, 9, 10, 12, 14, 16, 18, 20, 24, 28)), 3),
        ('CalibriB', 100, sweep_cases('Calibri', (9, 12, 14, 16, 18, 24), True), 3),
        ('Arial', 100, sweep_cases('Arial', (8, 9, 12, 14, 16, 18, 24)), 3),
        ('Times', 100, sweep_cases('Times New Roman', (10, 12, 14, 16, 20, 24)), 3),
        ('Black', 100, sweep_cases('Segoe UI Black', (9, 11, 14)), 3),
        ('Calibri75', 75, sweep_cases('Calibri', (11, 12, 14, 16, 18, 20, 24, 28)), 3),
        ('Calibri125', 125, sweep_cases('Calibri', (11, 12, 14, 16, 18, 20)), 3),
        ('Calibri85', 85, sweep_cases('Calibri', (11, 12, 14, 16, 18, 20, 24)), 3),
    ]),
}


class Styles:
    def __init__(self, normal):
        self.fonts = [normal]
        self.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
                    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>']

    def xf(self, align, font):
        if font not in self.fonts:
            self.fonts.append(font)
        xml = (f'<xf numFmtId="0" fontId="{self.fonts.index(font)}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
               f'<alignment {align}/></xf>')
        if xml not in self.xfs:
            self.xfs.append(xml)
        return self.xfs.index(xml)

    def xml(self):
        fonts = ''.join(f'<font>{"<b/>" if b else ""}<sz val="{s}"/><name val="{f}"/></font>' for f, s, b in self.fonts)
        return DECL + (
            f'<styleSheet xmlns="{MAIN}"><fonts count="{len(self.fonts)}">{fonts}</fonts>'
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FF000000"/><bgColor indexed="64"/></patternFill></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(self.xfs)}">{"".join(self.xfs)}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def col_name(i):
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[i]


def sheet(styles, zoom, cases, per_row, normal, selected):
    # The cases down the first column of cases, then the second ...: (the measure's order).
    rows_n = (len(cases) + per_row - 1) // per_row
    rows = []
    for k in range(rows_n):
        r = 2 * k + 2
        cells, height = '', max(15, normal[1] * 1.45)
        for j in range(per_row):
            i = k + j * rows_n
            if i >= len(cases):
                continue
            _, align, font, value, number = cases[i]
            height = max(height, font[1] * 1.45)
            a, b, c = (f'{col_name(4 * j + x)}{r}' for x in range(3))
            s = styles.xf(align, font)
            cells += (f'<c r="{a}" s="1"/>' + (f'<c r="{b}" s="{s}"><v>{value}</v></c>' if number else
                      f'<c r="{b}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{value}</t></is></c>') + f'<c r="{c}" s="1"/>')
        rows.append(f'<row r="{r - 1}" ht="6" customHeight="1"/>')
        rows.append(f'<row r="{r}" ht="{height:.2f}" customHeight="1">{cells}</row>')
    cols = ''
    for j in range(per_row):
        widths = (4.7109375, 24.7109375 if per_row < 3 else 18.7109375, 4.7109375, 2.7109375)
        for x, w in enumerate(widths):
            cols += f'<col min="{4 * j + x + 1}" max="{4 * j + x + 1}" width="{w}" customWidth="1"/>'
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if selected else ""}'
        f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{"".join(rows)}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def write(path, normal, specs):
    styles = Styles(normal)
    sheets = [(name, sheet(styles, zoom, cases, per, normal, i == 0)) for i, (name, zoom, cases, per) in enumerate(specs)]
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                      for i in range(len(sheets)))
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                                   + ''.join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (n, _) in enumerate(sheets))
                                   + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles.xml(),
    }
    for i, (_, xml) in enumerate(sheets):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)


if __name__ == '__main__':
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    for name in sys.argv[2:] or BOOKS:
        normal, specs = BOOKS[name]
        path = os.path.join(out, f'inset-{name}.xlsx')
        write(path, normal, specs)
        for sname, _, cases, _ in specs:
            print(f'{path}\t{sname}\t{" ".join(c[0] for c in cases)}')
