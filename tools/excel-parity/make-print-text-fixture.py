#!/usr/bin/env python3
"""Writes a synthetic workbook (not saved by Excel) for how Excel 2019 sets a cell's text when it prints (PDF export):
where it breaks wrapped text, how far a line stands from the cell's sides, how wide its text is against its column
(docs/excel-parity-tracker.md, "Cell text in print"). Checked with `print-excel-pdf.vbs` on the VM and
`pdftotext -bbox-layout` / `pdftoppm` of the PDF.

Each sheet is a font (SHEETS); column c (B..) has width WIDTHS[c] pixels at 100 %; per column: row 1 a label (the
width), row 2 the paragraph PARAGRAPH wrapped, at the top, in a row 300 pt high; row 3 "0000000000" left aligned,
row 4 the same right aligned, row 5 "nnnnnnnnnn" left aligned. Gridlines printed, no headings, 100 %.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-print-text-fixture.py OUT.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

PARAGRAPH = ('The quick brown fox jumps over the lazy dog while the monthly payment and the operating income '
             'of the property keep rising every single year')
SHEETS = [('Calibri', 11, False), ('Arial', 10, False), ('Calibri', 11, True), ('Times New Roman', 12, False)]
WIDTHS = list(range(40, 200, 6))


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def build(out):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>'] + [
        f'{"<b/>" if b else ""}<sz val="{s}"/><name val="{f}"/><family val="2"/>' for f, s, b in SHEETS]
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    for i in range(len(SHEETS)):
        fid = i + 1
        xfs += [f'<xf numFmtId="0" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>',
                f'<xf numFmtId="0" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left"/></xf>',
                f'<xf numFmtId="0" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>']
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    cols = '<cols>' + ''.join(f'<col min="{j + 2}" max="{j + 2}" width="{round(w / 7, 6)}" customWidth="1"/>'
                              for j, w in enumerate(WIDTHS)) + '</cols>'

    def s(t, st, r, c):
        return f'<c r="{col_letter(c)}{r}" s="{st}" t="inlineStr"><is><t>{escape(t)}</t></is></c>'
    sheets = []
    for i in range(len(SHEETS)):
        wrap, left, right = 1 + 3 * i, 2 + 3 * i, 3 + 3 * i
        cs = range(2, 2 + len(WIDTHS))
        rows = (f'<row r="1">' + ''.join(f'<c r="{col_letter(c)}1"><v>{WIDTHS[c - 2]}</v></c>' for c in cs) + '</row>'
                + '<row r="2" ht="300" customHeight="1">' + ''.join(s(PARAGRAPH, wrap, 2, c) for c in cs) + '</row>'
                + '<row r="3">' + ''.join(s('0000000000', left, 3, c) for c in cs) + '</row>'
                + '<row r="4">' + ''.join(s('0000000000', right, 4, c) for c in cs) + '</row>'
                + '<row r="5">' + ''.join(s('nnnnnnnnnn', left, 5, c) for c in cs) + '</row>')
        sheets.append((f'P{i + 1}', DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{rows}</sheetData>'
            '<printOptions gridLines="1"/>'
            '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
            '<pageSetup paperSize="1" orientation="landscape" scale="100"/></worksheet>')))
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
    build(sys.argv[1])
