#!/usr/bin/env python3
"""Writes synthetic workbooks (not saved by Excel) for where Excel 2019's PDF sets a cell's text in its cell and which
pages it prints (docs/excel-parity-tracker.md, "Print pages"): text of several fonts in rows of several heights at the
bottom, the middle and the top of their cells, left and right aligned, under several Normal fonts; and sheets whose
pages in the middle have nothing on them. Checked with `excel-print-pages.sh` on the VM.

    python3 tools/excel-parity/make-print-place-fixture.py OUT.xlsx [normal]

Sheets: "Place" - columns B..I in the fonts CELL_FONTS (B in the Normal font), "xH" in each; rows 2-10 of the heights
and vertical alignments ROWS (left aligned), row 11 right aligned; gridlines printed. "Blank" - "x" in A1 and K60
(four pages, two with nothing on them), a header "&P of &N"; "BlankGrid" the same with gridlines printed;
"BlankFill" - with a yellow fill in C55 too (a page with only a fill); "BlankBorder" - a border in M5 instead.

Deterministic: run it again and the bytes are the same.
"""
import sys
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

NORMALS = {'calibri11': ('Calibri', 11), 'segoeui10': ('Segoe UI', 10), 'timesnewroman12': ('Times New Roman', 12),
           'arial10': ('Arial', 10), 'tahoma10': ('Tahoma', 10)}
CELL_FONTS = [None, ('Calibri', 11), ('Arial', 10), ('Segoe UI', 10), ('Times New Roman', 12), ('Trebuchet MS', 10),
              ('Calibri', 20), ('Tahoma', 10)]
# (row, height in points or None for the default, vertical alignment)
# Formats of a number with a character's width after it (`_x`) or a literal one; the first without.
PADS = ['0', '0_)', '0_W', '0_.', '0\\ ', '0\\)', '"$"#,##0_);\\("$"#,##0\\)', '0_-', '0__']
ROWS = [(2, None, 'bottom'), (3, 30, 'bottom'), (4, 30, 'center'), (5, 30, 'top'), (6, 45, 'bottom'), (7, 45, 'center'),
        (8, 45, 'top'), (9, None, 'top'), (10, None, 'center')]


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def build(out, normal):
    face, size = NORMALS[normal]
    fonts = [f'<sz val="{size}"/><name val="{face}"/><family val="2"/>'] + [
        f'<sz val="{s}"/><name val="{f}"/><family val="2"/>' for f, s in [x for x in CELL_FONTS if x]]
    # xf: 0 Normal; then per (font, vertical) a left-aligned style, per font a right-aligned bottom one; a fill; a border.
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for i, f in enumerate(CELL_FONTS):
        fid = 0 if f is None else CELL_FONTS[1:].index(f) + 1
        for v in ('bottom', 'center', 'top'):
            style[(i, v)] = len(xfs)
            xfs.append(f'<xf numFmtId="0" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
                       f'<alignment vertical="{v}"/></xf>')
        style[(i, 'right')] = len(xfs)
        xfs.append(f'<xf numFmtId="0" fontId="{fid}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
                   '<alignment horizontal="right"/></xf>')
    # Right-aligned numbers in formats ending in a width of a character (PADS: 164 ...).
    pad_xf = len(xfs)
    for i in range(len(PADS)):
        xfs.append(f'<xf numFmtId="{164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
    fill_xf = len(xfs)
    xfs.append('<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>')
    border_xf = len(xfs)
    xfs.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>')
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(PADS)}">' + ''.join(f'<numFmt numFmtId="{164 + i}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>' for i, code in enumerate(PADS)) + '</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right>'
        '<top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    def cell(r, c, text, st):
        sa = f' s="{st}"' if st else ''
        if text is None:
            return f'<c r="{col_letter(c)}{r}"{sa}/>'
        return f'<c r="{col_letter(c)}{r}"{sa} t="inlineStr"><is><t>{escape(text)}</t></is></c>'

    sheets = []
    # Place
    rows = [f'<row r="1">' + ''.join(cell(1, c + 2, (f'{f[0]} {f[1]}' if f else 'Normal'), 0) for c, f in enumerate(CELL_FONTS)) + '</row>']
    for r, h, v in ROWS:
        ht = f' ht="{h}" customHeight="1"' if h else ''
        rows.append(f'<row r="{r}"{ht}>' + cell(r, 1, f'{h or "default"} {v}', 0)
                    + ''.join(cell(r, c + 2, 'xH', style[(c, v)]) for c in range(len(CELL_FONTS))) + '</row>')
    rows.append('<row r="11" ht="30" customHeight="1">' + cell(11, 1, '30 right', 0)
                + ''.join(cell(11, c + 2, 'xH', style[(c, 'right')]) for c in range(len(CELL_FONTS))) + '</row>')
    cols = '<cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="9" width="12" customWidth="1"/></cols>'
    sheets.append(('Place', cols + f'<sheetData>{"".join(rows)}</sheetData>', '<printOptions gridLines="1"/>', ''))

    def blank(marks, extra_cells, grid):
        rows = {}
        for r, c in marks:
            rows.setdefault(r, []).append((c, cell(r, c, 'x', 0)))
        for r, c, st in extra_cells:
            rows.setdefault(r, []).append((c, cell(r, c, None, st)))
        xml = ''.join(f'<row r="{r}">' + ''.join(x for _, x in sorted(rows[r])) + '</row>' for r in sorted(rows))
        return f'<sheetData>{xml}</sheetData>', '<printOptions gridLines="1"/>' if grid else '', '<headerFooter><oddHeader>&amp;CPage &amp;P of &amp;N</oddHeader></headerFooter>'
    corner = [(1, 1), (60, 11)]
    for name, marks, extra, grid in [('Blank', corner, [], False), ('BlankGrid', corner, [], True),
                                     ('BlankFill', corner, [(55, 3, fill_xf)], False), ('BlankBorder', corner, [(5, 13, border_xf)], False),
                                     # (the last page - down, then over - with nothing on it; two such pages last)
                                     ('BlankLast', [(1, 1), (1, 11), (60, 1)], [], False),
                                     ('BlankLastTwo', [(1, 1), (1, 11), (110, 1)], [], False),
                                     ('BlankLastGrid', [(1, 1), (1, 11), (60, 1)], [], True)]:
        data, options, hf = blank(marks, extra, grid)
        sheets.append((name, data, options, hf))
    # Numbers right aligned (as numbers are) in the formats PADS, a column each, rows of 123 and 4567.
    prow = ''.join(f'<row r="{r}">' + ''.join(f'<c r="{col_letter(c + 2)}{r}" s="{pad_xf + c}"><v>{v}</v></c>' for c in range(len(PADS))) + '</row>'
                   for r, v in [(2, 123), (3, 4567)])
    sheets.append(('Pads', cols + f'<sheetData><row r="1">' + ''.join(cell(1, c + 2, code, 0) for c, code in enumerate(PADS)) + f'</row>{prow}</sheetData>',
                   '<printOptions gridLines="1"/>', ''))
    # Text running over to the columns of a last page whose rows have nothing (acre-property's Summary).
    data, options, hf = blank([(60, 1)], [], False)
    long = ' '.join(['The quick brown fox jumps over the lazy dog'] * 3)
    data = data.replace('<sheetData>', f'<sheetData><row r="1">{cell(1, 1, long, 0)}</row>')
    sheets.append(('BlankOverflow', data, options, hf))

    parts = []
    for i, (name, data, options, hf) in enumerate(sheets):
        parts.append(DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView{" tabSelected=\"1\"" if i == 0 else ""} workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/>{data}{options}'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'<pageSetup paperSize="1" orientation="portrait"/>{hf}</worksheet>'))
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
                                   + ''.join(f'<sheet name="{s[0]}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, s in enumerate(sheets))
                                   + '</sheets></workbook>'),
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
        for k in NORMALS:
            build(sys.argv[1].replace('.xlsx', f'-{k}.xlsx'), k)
