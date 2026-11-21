#!/usr/bin/env python3
"""Writes synthetic-error-checks.xlsx (not saved by Excel): cells for the background error checks Excel 2019 runs
by default and marks with a green triangle (checked in Excel 2019 on the VM: `error-probe.vbs` reads
`Range.Errors(n).Value` of every cell, `shoot-sheets.ps1` the picture; tools/xlsx-browser check
`error-triangles-excel-parity`, tests/office-formats/xlsx-bridge.fsx "error checks:").

- "Checks": numbers stored as text (A), a formula omitting the cell next to its range (B), formulas inconsistent
  with their neighbours in a row (row 10) and a column (J), formulas evaluating to errors and a typed error (K),
  a text date of a two-digit year (L).
- "Ignored": the same cells with `<ignoredErrors>` saying so for each kind.

Deterministic: run it again and the bytes are the same.

    python3 tools/excel-parity/make-error-check-fixtures.py
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.join(os.path.dirname(os.path.dirname(HERE)), 'tests', 'test-documents-excel')
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'


def esc(t):
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def text(ref, value, style=0):
    return f'<c r="{ref}" t="inlineStr"' + (f' s="{style}"' if style else '') + f'><is><t xml:space="preserve">{esc(value)}</t></is></c>'


def num(ref, value):
    return f'<c r="{ref}"><v>{value}</v></c>'


def formula(ref, f, value, kind=None):
    t = f' t="{kind}"' if kind else ''
    return f'<c r="{ref}"{t}><f>{esc(f)}</f><v>{esc(value)}</v></c>'


def cells():
    out = {}

    def put(ref, xml):
        row = int(''.join(ch for ch in ref if ch.isdigit()))
        out.setdefault(row, []).append(xml)

    # A: text that reads as a number, and text that does not.
    for i, value in enumerate(['123', '1,000', 'abc', ' 12', '12%', '1e3', '$5', '-4.5', '(7)', '0012']):
        put(f'A{i + 1}', text(f'A{i + 1}', value))
    put('A11', text('A11', '123', 1))     # quotePrefix
    # B: numbers 1 .. 4, a SUM that leaves B4 out, one that does not, one of a row that leaves D20 out.
    for i in range(4):
        put(f'B{i + 1}', num(f'B{i + 1}', i + 1))
    put('B5', formula('B5', 'SUM(B1:B3)', 6))
    put('B6', formula('B6', 'SUM(B1:B4)', 10))
    # Row 9: numbers; row 10: =x9+1 but F10 =F9*5.
    for col in 'CDEFGHI':
        put(f'{col}9', num(f'{col}9', 2))
    for col in 'CDEFGHI':
        put(f'{col}10', formula(f'{col}10', f'{col}9*5' if col == 'F' else f'{col}9+1', 10 if col == 'F' else 3))
    # Column J: =I1*2 .. but J3 =I3+7.
    for r in range(1, 7):
        put(f'I{r}', num(f'I{r}', r))
        put(f'J{r}', formula(f'J{r}', f'I{r}+7' if r == 3 else f'I{r}*2', r + 7 if r == 3 else r * 2))
    # K: errors.
    put('K1', formula('K1', '1/0', '#DIV/0!', 'e'))
    put('K2', formula('K2', 'NA()', '#N/A', 'e'))
    put('K3', '<c r="K3" t="e"><v>#REF!</v></c>')
    put('K4', formula('K4', 'IFERROR(1/0,0)', 0))
    # L: a date as text with a two-digit year.
    put('L1', text('L1', '1/1/99'))
    put('L2', text('L2', 'Jan-99'))
    # Row 20: numbers C20..F20, G20 =SUM(C20:E20) omits F20 (beside its range, not beside the cell).
    for col in 'CDEF':
        put(f'{col}20', num(f'{col}20', 5))
    put('G20', formula('G20', 'SUM(C20:E20)', 15))
    # Row 22: numbers C22..F22, G22 =SUM(C22:F22), H22 =SUM(C22:F22) (the formula next to it left out).
    for col in 'CDEF':
        put(f'{col}22', num(f'{col}22', 1))
    put('G22', formula('G22', 'SUM(C22:F22)', 4))
    put('H22', formula('H22', 'SUM(C22:F22)', 4))
    return out


def sheet(extra=''):
    rows = ''.join(f'<row r="{r}">' + ''.join(sorted(cs, key=lambda c: (len(c.split('"')[1]), c.split('"')[1]))) + '</row>' for r, cs in sorted(cells().items()))
    return (DECL + f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
            f'<sheetFormatPr defaultRowHeight="15"/><sheetData>{rows}</sheetData>{extra}</worksheet>')


def build():
    ignored = ('<ignoredErrors><ignoredError sqref="A1:A11 L1:L2" numberStoredAsText="1" twoDigitTextYear="1"/>'
               '<ignoredError sqref="B5 G20 H22" formulaRange="1"/><ignoredError sqref="F10 J3" formula="1"/>'
               '<ignoredError sqref="K1:K4" evalError="1"/></ignoredErrors>')
    sheets = [('Checks', sheet()), ('Ignored', sheet(ignored))]
    styles = (DECL + f'<styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/></cellXfs>'
              '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>'
            + ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>' for i in range(len(sheets)))
            + '</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
                                   + ''.join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, (n, _) in enumerate(sheets))
                                   + '</sheets><calcPr calcId="191029"/></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (f'<Relationships xmlns="{PKG}">'
                                              + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
                                              + f'<Relationship Id="rId{len(sheets) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheets):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    path = os.path.join(OUTDIR, 'synthetic-error-checks.xlsx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, body in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, body)
    print(path)


if __name__ == '__main__':
    build()
