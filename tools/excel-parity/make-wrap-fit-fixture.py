#!/usr/bin/env python3
"""Writes a synthetic workbook (not saved by Excel) for how many lines Excel 2019 fits a row of wrapped text to when it
opens a file (and when AutoFit fits it): two-word texts in columns of every width (pixels at 100 %) of a range, one
cell a row, so a row's height says whether the text took one line or two at that width
(`probe-wrap-fit.vbs` reads `Rows(r).Height` on opening, after `Rows.AutoFit` at the sheet's zoom and at 100 %;
docs/excel-parity-tracker.md, "Wrapped rows on opening").

Each sheet is a font, size, style and zoom (SHEETS); row 1 + k * len(WIDTHS) + j holds text k in the column of width
WIDTHS[j] (column j + 2). A sidecar `.cases.tsv` lists sheet, row, font, size, bold, zoom, text, width.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-wrap-fit-fixture.py OUT.xlsx [CASES.tsv]
"""
import os
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

TEXTS = ['Cost Unit', 'The quick', 'brown fox', 'nnnn nnnn', '0000 0000', 'iiii llll', 'WWWW MMMM',
         'Releasing Cost', 'abc defghijklm', 'Monthly Payment', 'Total Revenue', 'x xxxxxxxxxxxxxx',
         'aaaaaaaaaaaa b', 'Operating Income']
# (face, size, bold, zoom)
SWEEP = ([('Calibri', 11, False, z) for z in (100, 85, 75, 120, 150, 200)]
         + [('Calibri', 11, True, z) for z in (100, 85, 150)]
         + [('Arial', 10, False, z) for z in (100, 85, 75, 120, 150, 200)]
         + [('Arial', 11, False, 100), ('Times New Roman', 11, False, 100), ('Verdana', 10, False, 100),
            ('Calibri', 14, False, 100), ('Calibri', 14, False, 85), ('Segoe UI', 9, False, 100)])
WIDTHS = list(range(20, 201))


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def build(out, sheets_spec, texts, widths, windows=None):
    """windows: per sheet, per text, the (lo, hi) widths it is put in (else every width)."""
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for face, size, bold, _ in sheets_spec:
        fonts.append(f'{"<b/>" if bold else ""}<sz val="{size}"/><name val="{face}"/><family val="2"/>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'] + [
        f'<xf numFmtId="0" fontId="{i + 1}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
        '<alignment wrapText="1"/></xf>' for i in range(len(sheets_spec))]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    cols = '<cols>' + ''.join(f'<col min="{j + 2}" max="{j + 2}" width="{round(w / 7, 6)}" customWidth="1"/>'
                              for j, w in enumerate(widths)) + '</cols>'
    cases = []
    sheets = []
    for i, (face, size, bold, zoom) in enumerate(sheets_spec):
        name = f'S{i + 1}'
        rows = []
        r = 0
        for k, text in enumerate(texts[i] if windows else texts):
            for j, w in enumerate(widths):
                if windows and not (windows[i][k][0] <= w <= windows[i][k][1]):
                    continue
                r += 1
                rows.append(f'<row r="{r}"><c r="{col_letter(j + 2)}{r}" s="{i + 1}" t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c></row>')
                cases.append(f'{name}\t{r}\t{face}\t{size}\t{int(bold)}\t{zoom}\t{text}\t{w}')
        sheets.append((name, DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""}'
            f' zoomScale="{zoom}" zoomScaleNormal="{zoom}" workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{"".join(rows)}</sheetData>'
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
    with open(os.path.splitext(out)[0] + '.cases.tsv', 'w') as f:
        f.write('\n'.join(cases) + '\n')
    print(out)


def from_cases(path):
    """A sweep from a file of lines `face<TAB>size<TAB>bold<TAB>zoom<TAB>text<TAB>lo<TAB>hi`, a sheet per font."""
    sheets, texts, windows = [], [], []
    for line in open(path, encoding='utf-8'):
        face, size, bold, zoom, text, lo, hi = line.rstrip('\n').split('\t')
        key = (face, float(size) if '.' in size else int(size), bold == '1', int(zoom))
        if key not in sheets:
            sheets.append(key); texts.append([]); windows.append([])
        i = sheets.index(key)
        texts[i].append(text); windows[i].append((int(lo), int(hi)))
    return sheets, texts, windows


if __name__ == '__main__':
    if len(sys.argv) > 2:
        sheets, texts, windows = from_cases(sys.argv[2])
        build(sys.argv[1], sheets, texts, list(range(min(w[0] for ws in windows for w in ws), max(w[1] for ws in windows for w in ws) + 1)), windows)
    else:
        build(sys.argv[1], SWEEP, TEXTS, WIDTHS)
