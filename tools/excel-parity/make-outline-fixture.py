#!/usr/bin/env python3
"""Writes tests/test-documents-excel/synthetic-outline.xlsx, a synthetic workbook (not saved by Excel): grouped rows and
columns (Data > Group), for the outline's strips beside and above the headers - the level buttons, the bars and the
+ / - buttons - against Excel 2019's (pictures with `shoot-sheets.ps1`; docs/excel-parity-tracker.md, "Outline").

"Rows": rows 2-9 at level 1 and 4-6 at level 2 (summary rows below: 10, 7), rows 12-14 a collapsed group (13-14 hidden,
15 `collapsed`). "Cols": columns B-F at level 1, C-D at level 2. "Both": rows 2-6 and columns B-D at level 1, rows 3-4 at
level 2 and 3; "Above": the rows' summary above (`summaryBelow="0"`), the columns' to the left (`summaryRight="0"`).

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-outline-fixture.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.normpath(os.path.join(HERE, '..', '..'))
OUT = os.path.join(REPO, 'tests', 'test-documents-excel', 'synthetic-outline.xlsx')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'


def col_letter(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def sheet(row_levels, col_levels, hidden_rows=(), collapsed_rows=(), above=False, selected=False):
    rows_max = max(list(row_levels) + [16])
    cols = ''.join(f'<col min="{c}" max="{c}" width="9.140625" outlineLevel="{lvl}"/>' for c, lvl in sorted(col_levels.items()))
    body = []
    for r in range(1, rows_max + 1):
        attrs = ''
        if r in row_levels:
            attrs += f' outlineLevel="{row_levels[r]}"'
        if r in hidden_rows:
            attrs += ' hidden="1"'
        if r in collapsed_rows:
            attrs += ' collapsed="1"'
        cells = ''.join(f'<c r="{col_letter(c)}{r}"><v>{r * 10 + c}</v></c>' for c in range(1, 9))
        body.append(f'<row r="{r}"{attrs}>{cells}</row>')
    max_row = max(list(row_levels.values()) + [0])
    max_col = max(list(col_levels.values()) + [0])
    summary = ' summaryBelow="0" summaryRight="0"' if above else ''
    return DECL + (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetPr><outlinePr{summary}/></sheetPr>'
        f'<sheetViews><sheetView{" tabSelected=\"1\"" if selected else ""} workbookViewId="0"/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15" outlineLevelRow="{max_row}" outlineLevelCol="{max_col}"/>'
        + (f'<cols>{cols}</cols>' if cols else '') +
        f'<sheetData>{"".join(body)}</sheetData>'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def build():
    sheets = [
        ('Rows', sheet({**{r: 1 for r in range(2, 10)}, **{r: 2 for r in range(4, 7)}, 13: 1, 14: 1}, {},
                       hidden_rows=(13, 14), collapsed_rows=(15,), selected=True)),
        ('Cols', sheet({}, {**{c: 1 for c in range(2, 7)}, 3: 2, 4: 2})),
        ('Both', sheet({2: 1, 3: 3, 4: 2, 5: 1, 6: 1}, {2: 1, 3: 1, 4: 1})),
        ('Above', sheet({**{r: 1 for r in range(3, 8)}, 5: 2}, {c: 1 for c in range(3, 6)}, above=True)),
    ]
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
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
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    print(OUT)


if __name__ == '__main__':
    build()
