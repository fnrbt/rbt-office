#!/usr/bin/env python3
"""Fixtures for a numbered list across section breaks (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/list-sections.py [out dir]   (default tests/word/fixtures/list-sections)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/list-sections \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/list-sections --refresh-word
    python3 tools/word-parity/fixtures/list-sections.py --numbers tests/word/fixtures/list-sections

--numbers reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/numbers.tsv:
every label's fixture, page and the list number before it. tests/word/list-sections.fsx checks ours.

The showcase handbook's numbered steps (List Number, one list through the document) are 10, 11, 12 in
Word after a landscape section, 1, 2, 3 in ours. Each fixture: Calibri 11pt; a01 - a03 in one decimal
list (numId 1), a section break, b01 - b03 in the same list, another break, c01 (numId 1) and d01 in a
second num of the same abstract list (numId 2). Varies:
  next / cont / land   the breaks: next page, continuous, next page to a landscape section and back
  c14 / c15            compatibility mode
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
NUMBERING = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>
<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>'''
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'''
PORTRAIT = '<w:pgSz w:w="11907" w:h="16839"/>'
LANDSCAPE = '<w:pgSz w:w="16839" w:h="11907" w:orient="landscape"/>'
MAR = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


def item(label, num=1):
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr></w:pPr><w:r><w:t>{label} item</w:t></w:r></w:p>'


def sect(kind, size):
    t = '<w:type w:val="continuous"/>' if kind == 'cont' else ''
    return f'<w:p><w:pPr><w:sectPr>{t}{size}{MAR}</w:sectPr></w:pPr></w:p>'


def body(kind):
    middle = LANDSCAPE if kind == 'land' else PORTRAIT
    out = ''.join(item(f'a0{i}') for i in (1, 2, 3)) + sect(kind, PORTRAIT)
    out += ''.join(item(f'b0{i}') for i in (1, 2, 3)) + sect(kind, middle)
    out += item('c01') + item('d01', 2)
    t = '<w:type w:val="continuous"/>' if kind == 'cont' else ''
    return out + f'<w:sectPr>{t}{PORTRAIT}{MAR}</w:sectPr>'


def write(path, xml, mode):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>{xml}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in (('[Content_Types].xml', TYPES), ('_rels/.rels', RELS), ('word/_rels/document.xml.rels', DOC_RELS),
                           ('word/document.xml', doc), ('word/styles.xml', STYLES), ('word/settings.xml', settings(mode)),
                           ('word/numbering.xml', NUMBERING)):
            z.writestr(zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0)), data, zipfile.ZIP_DEFLATED)


def docs():
    return {f'{kind}-c{mode}': (body(kind), mode) for mode in (14, 15) for kind in ('next', 'cont', 'land')}


def numbers(folder):
    rows = []
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(folder, 'word', name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        out = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for page, text in enumerate(out.split('\f'), 1):
            for m in re.finditer(r'(\d+)\.\s+([a-d]\d\d) item', text):
                rows.append(f'{name}\t{page}\t{m.group(2)}\t{m.group(1)}')
    with open(os.path.join(folder, 'word', 'numbers.tsv'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} items')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--numbers':
        numbers(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/list-sections')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    ds = docs()
    for name, (xml, mode) in ds.items():
        write(os.path.join(out_dir, name + '.docx'), xml, mode)
    print(f'{len(ds)} fixtures in {out_dir}')
