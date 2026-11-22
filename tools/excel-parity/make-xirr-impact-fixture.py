#!/usr/bin/env python3
"""Generate a synthetic XIRR workbook for eval-file-on-vm.sh and xirr-impact-engine.fsx.

    python3 tools/excel-parity/make-xirr-impact-fixture.py /tmp/xirr-impact.xlsx

Two dated flows per row, no private workbook inputs. Dates are Excel serials; the
one-year cases are exactly 365 days apart. Values/thresholds are decimal inputs.

Excel 2019 and engine observations: xirr-impact-results.tsv. Excel values are
the 15-significant-digit values from eval-file.vbs, not exact binary encodings;
that tool also emits the exact binary values when a deeper comparison is needed.
Raw and rounded comparisons use each row's hurdle; the monthly rows use 10%.
The month-12 explicit-guess rate is F26; all other rates are in column G.
    tools/excel-parity/eval-file-on-vm.sh /tmp/xirr-impact.xlsx /tmp/xirr-impact-excel.tsv
    dotnet fsi tools/excel-parity/xirr-impact-engine.fsx /tmp/xirr-impact.xlsx
"""
import sys
import zipfile
from xml.sax.saxutils import escape

# name, initial, terminal, day interval, guess (None means omitted), hurdle
CASES = [
    ('positive-default', -1000, 1100, 365, None, '0.1'),
    ('positive-guess-0.5', -1000, 1100, 365, '0.5', '0.1'),
    ('positive-guess-negative', -1000, 1100, 365, '-0.1', '0.1'),
    ('positive-scale-small', -1, 1.1, 365, None, '0.1'),
    ('positive-scale-large', -1000000000, 1100000000, 365, None, '0.1'),
    ('reverse-sign', 1000, -1100, 365, None, '0.1'),
    ('negative-default', -1000, 900, 365, None, '-0.1'),
    ('negative-guess-negative', -1000, 900, 365, '-0.1', '-0.1'),
    ('rounding-boundary', -1000, '1100.0005', 365, None, '0.1000005'),
    ('rounding-boundary-guess', -1000, '1100.0005', 365, '0.5', '0.1000005'),
    ('180-day-holdout', -1000, 1050, 180, None, '0.1'),
]

def cell(ref, value, kind='n', style=''):
    s = f' s="{style}"' if style else ''
    if kind == 's':
        return f'<c r="{ref}" t="inlineStr"{s}><is><t>{escape(str(value))}</t></is></c>'
    if kind == 'f':
        return f'<c r="{ref}"{s}><f>{escape(str(value))}</f></c>'
    return f'<c r="{ref}"{s}><v>{value}</v></c>'

rows = []
headers = ['case', 'initial', 'terminal', 'start', 'end', 'guess', 'XIRR', 'hurdle', 'ROUND(rate,6)', 'rate>=hurdle', 'rounded>=rounded hurdle', 'decision']
rows.append('<row r="1">' + ''.join(cell(f'{chr(65+i)}1', h, 's') for i, h in enumerate(headers)) + '</row>')
for r, (name, initial, terminal, days, guess, hurdle) in enumerate(CASES, 2):
    cells = [cell(f'A{r}', name, 's'), cell(f'B{r}', initial), cell(f'C{r}', terminal),
             cell(f'D{r}', 43831), cell(f'E{r}', 43831 + days)]
    if guess is not None:
        cells.append(cell(f'F{r}', guess))
    arg = f',F{r}' if guess is not None else ''
    cells += [cell(f'G{r}', f'XIRR(B{r}:C{r},D{r}:E{r}{arg})', 'f', '1'),
              cell(f'H{r}', hurdle), cell(f'I{r}', f'ROUND(G{r},6)', 'f'),
              cell(f'J{r}', f'G{r}>=H{r}', 'f'),
              cell(f'K{r}', f'I{r}>=ROUND(H{r},6)', 'f'),
              cell(f'L{r}', f'IF(J{r},"accept","reject")', 'f')]
    rows.append(f'<row r="{r}">' + ''.join(cells) + '</row>')

# A synthetic monthly waterfall: small distributions followed by principal.
# Rolling XIRR reports use IFERROR/ROUND as in the traced reporting path.
for r in range(14, 27):
    month = r - 14
    date = [43831, 43862, 43891, 43922, 43952, 43983, 44013,
            44044, 44075, 44105, 44136, 44166, 44197][month]
    flow = -1000 if month == 0 else (1050 if month == 12 else 10)
    cells = [cell(f'A{r}', f'month-{month:02}', 's'), cell(f'B{r}', flow),
             cell(f'C{r}', date)]
    if month == 12:
        cells.append(cell(f'F{r}', 'XIRR(B$14:B26,C$14:C26,0.1)', 'f', '1'))
    if month:
        cells += [cell(f'G{r}', f'XIRR(B$14:B{r},C$14:C{r})', 'f', '1'),
                  cell(f'I{r}', f'IFERROR(ROUND(G{r},6),\"-\")', 'f'),
                  cell(f'J{r}', f'IFERROR(G{r}>=0.1,FALSE)', 'f'),
                  cell(f'K{r}', f'IFERROR(I{r}>=0.1,FALSE)', 'f'),
                  cell(f'L{r}', f'IF(J{r},\"accept\",\"reject\")', 'f')]
    rows.append(f'<row r=\"{r}\">' + ''.join(cells) + '</row>')
main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
pkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
files = {
 '[Content_Types].xml': f'<Types xmlns="{pkg}/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ct}.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/></Types>'.replace('schemas.openxmlformats.org/package/2006/relationships/content-types', 'schemas.openxmlformats.org/package/2006/content-types'),
 '_rels/.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
 'xl/workbook.xml': f'<workbook xmlns="{main}" xmlns:r="{rel}"><bookViews><workbookView/></bookViews><sheets><sheet name="Impact" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>',
 'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{rel}/styles" Target="styles.xml"/></Relationships>',
 'xl/styles.xml': f'<styleSheet xmlns="{main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
 'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{main}" xmlns:r="{rel}"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="12" width="20" customWidth="1"/></cols><sheetData>{"".join(rows)}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>',
}
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in files.items():
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text)
