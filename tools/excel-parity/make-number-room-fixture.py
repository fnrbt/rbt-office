#!/usr/bin/env python3
"""Writes a synthetic workbook (not saved by Excel) for the room a number has in its cell in italic and upright fonts:
the widths at which Excel 2019 shows a formatted number rather than `#`s (`probe-display.vbs` reads `Range.Text` at
each sheet's zoom; docs/excel-parity-tracker.md, "Numbers in italic").

Each sheet is a font, style and zoom (SHEETS); columns B.. are WIDTHS pixels wide at 100 %; row 1 holds -1 in `0.00%`
(-100.00%), row 2 0.5 in `0.00%` (50.00%), row 3 1234567 in `#,##0` (1,234,567), row 4 -1234.5 in `#,##0.00`.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-number-room-fixture.py OUT.xlsx [sizes]
"""
import sys
import zipfile

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

SHEETS = [(face, 14 if face == 'Aptos Narrow' else 11, italic, bold, zoom)
          for face in ('Aptos Narrow', 'Calibri', 'Arial')
          for italic, bold in ((False, False), (True, False), (True, True), (False, True))
          for zoom in (100, 57)]
# More sizes of italic (`sizes`): Calibri and Arial italic at 8 - 22 pt, and in Small Fonts at 50 and 67 %.
SIZES = ([(face, size, True, False, 100) for face in ('Calibri', 'Arial', 'Times New Roman') for size in (8, 10, 12, 14, 16, 18, 22)]
         + [(face, 11, True, False, zoom) for face in ('Calibri', 'Arial') for zoom in (50, 67)])
WIDTHS = list(range(20, 121))
VALUES = [(-1, 10), (0.5, 10), (1234567, 3), (-1234.5, 4)]


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def build(out, sheets_spec):
    global SHEETS
    SHEETS = sheets_spec
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>'] + [
        f'{"<b/>" if b else ""}{"<i/>" if i else ""}<sz val="{s}"/><name val="{f}"/><family val="2"/>' for f, s, i, b, _ in SHEETS]
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    for k in range(len(SHEETS)):
        for _, fmt in VALUES:
            xfs.append(f'<xf numFmtId="{fmt}" fontId="{k + 1}" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    cols = '<cols>' + ''.join(f'<col min="{j + 2}" max="{j + 2}" width="{round(w / 7, 6)}" customWidth="1"/>' for j, w in enumerate(WIDTHS)) + '</cols>'
    sheets = []
    for k, (face, size, italic, bold, zoom) in enumerate(SHEETS):
        rows = ''.join(f'<row r="{r + 1}">' + ''.join(f'<c r="{col_letter(j + 2)}{r + 1}" s="{1 + k * len(VALUES) + r}"><v>{v}</v></c>' for j in range(len(WIDTHS))) + '</row>'
                       for r, (v, _) in enumerate(VALUES))
        name = f'{face.split()[0][:6]}{"I" if italic else ""}{"B" if bold else ""}{zoom}' + (f'_{size}' if size not in (11, 14) else '')
        sheets.append((name, DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if k == 0 else ""}'
            f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{rows}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')))
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
                                   + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (name, _) in enumerate(sheets))
                                   + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheets):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(out)


if __name__ == '__main__':
    build(sys.argv[1], SIZES if len(sys.argv) > 2 and sys.argv[2] == 'sizes' else SHEETS)
