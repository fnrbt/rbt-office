#!/usr/bin/env python3
"""A workbook of cells with furigana (`<rPh>` runs and `<phoneticPr>`), for PHONETIC.

    python3 tools/excel-parity/make-phonetic-fixture.py tests/test-documents-excel/synthetic-phonetic.xlsx

Column A holds the texts (shared strings with and without readings, a rich text whose runs carry
readings, a reading that covers part of the text, an inline string, a number, an empty cell); column B
holds =PHONETIC(A<n>), and B20 on are PHONETIC of ranges. Opened in Excel on the VM
(eval-file-on-vm.sh ... none / full) it gives what Excel answers; phonetic-excel.tsv keeps that.
"""
import sys
import zipfile

strings = [
    # 0: a whole-text reading
    '<si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="1"/></si>',
    # 1: two readings, text between them without one
    '<si><t>東京都港区</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><rPh sb="2" eb="3"><t>ト</t></rPh>'
    '<rPh sb="3" eb="4"><t>ミナト</t></rPh><rPh sb="4" eb="5"><t>ク</t></rPh><phoneticPr fontId="1"/></si>',
    # 2: a reading over part of the text only
    '<si><t>山田さん</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><phoneticPr fontId="1"/></si>',
    # 3: no reading, phoneticPr only
    '<si><t>abc</t><phoneticPr fontId="1"/></si>',
    # 4: no phonetic data at all
    '<si><t>plain</t></si>',
    # 5: rich text runs, readings after them
    '<si><r><t>日本</t></r><r><rPr><b/></rPr><t>語</t></r><rPh sb="0" eb="2"><t>ニホン</t></rPh>'
    '<rPh sb="2" eb="3"><t>ゴ</t></rPh><phoneticPr fontId="1" type="Hiragana"/></si>',
    # 6: hiragana readings, type Hiragana
    '<si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1" type="Hiragana"/></si>',
    # 7: a reading in the middle
    '<si><t>私の名前</t><rPh sb="0" eb="1"><t>ワタシ</t></rPh><rPh sb="2" eb="4"><t>ナマエ</t></rPh><phoneticPr fontId="1"/></si>',
    # 8: phoneticPr type noConversion, half-width reading
    '<si><t>半角</t><rPh sb="0" eb="2"><t>ﾊﾝｶｸ</t></rPh><phoneticPr fontId="1" type="halfwidthKatakana"/></si>',
]
# (row, kind, value)
cells = [(1, 's', 0), (2, 's', 1), (3, 's', 2), (4, 's', 3), (5, 's', 4), (6, 's', 5), (7, 's', 6), (8, 's', 7),
         (9, 's', 8), (10, 'n', 123), (11, 'b', 1), (13, 'str', '漢字'), (14, 'i', '漢字')]
formulas = [('B%d' % r, 'PHONETIC(A%d)' % r) for r in range(1, 15)]
formulas += [('B20', 'PHONETIC(A1:A2)'), ('B21', 'PHONETIC(A1:A9)'), ('B22', 'PHONETIC(A4:A5)'),
             ('B23', 'PHONETIC(A12:A14)'), ('B24', 'PHONETIC(A10:A11)'), ('B25', 'PHONETIC(A2:A3)'),
             ('B26', 'PHONETIC(A1:C1)'), ('B27', 'LEN(PHONETIC(A2))'), ('B28', 'PHONETIC(A3:A1)'),
             ('B29', 'PHONETIC((A1,A2))'), ('B30', 'PHONETIC(A12)'), ('B31', 'PHONETIC(A5:A6)')]

rows = {}
for r, kind, v in cells:
    ref = 'A%d' % r
    if kind == 's':
        x = '<c r="%s" t="s"><v>%d</v></c>' % (ref, v)
    elif kind == 'n':
        x = '<c r="%s"><v>%s</v></c>' % (ref, v)
    elif kind == 'b':
        x = '<c r="%s" t="b"><v>%d</v></c>' % (ref, v)
    elif kind == 'str':
        x = '<c r="%s" t="str"><f>"漢"&amp;"字"</f><v>%s</v></c>' % (ref, v)
    else:
        x = '<c r="%s" t="inlineStr"><is><t>%s</t><rPh sb="0" eb="2"><t>カンジ</t></rPh></is></c>' % (ref, v)
    rows.setdefault(r, []).append(x)
for ref, f in formulas:
    r = int(ref[1:])
    rows.setdefault(r, []).append('<c r="%s" t="e"><f>%s</f><v>#N/A</v></c>' % (ref, f.replace('&', '&amp;')))
sheet = ''.join('<row r="%d">%s</row>' % (r, ''.join(rows[r])) for r in sorted(rows))

files = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>',
    'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>',
    'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="6"/><name val="Yu Gothic"/><charset val="128"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '</styleSheet>',
    'xl/sharedStrings.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">%s</sst>'
    % (len(strings), len(strings), ''.join(strings)),
    'xl/worksheets/sheet1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>%s</sheetData>'
    '<phoneticPr fontId="1"/></worksheet>' % sheet,
}
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in files.items():
        z.writestr(name, text.encode('utf-8'))
