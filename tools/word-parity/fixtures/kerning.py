#!/usr/bin/env python3
"""Fixtures for whether Word kerns text, and when (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/kerning.py [out dir]   (default tests/word/fixtures/kerning)

then, to ask Word (winvm) and read the widths it gave each word:

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/kerning \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/kerning --refresh-word
    python3 tools/word-parity/fixtures/kerning.py --widths tests/word/fixtures/kerning

--widths reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/widths.tsv:
fixture, font, size, word, the word's width in Word's PDF (pdftotext's box: from its first glyph's
left to its last glyph's advance). tests/word/kerning.fsx checks our widths against it.

Every fixture is one paragraph per (font, word): each word alone on its line, left aligned, in
Calibri, Arial, Times New Roman and Cambria. The words are made of pairs the fonts kern strongly
(AV, WA, To, LT, Yo) and of letters Calibri makes ligatures of (fi, ffi, fl). What varies:

  kern-none[-<size>]     no w:kern (11pt; -16: 16pt)
  kern-2[-<size>]        w:kern w:val="2" on every run (kern at 1pt and up)
  kern-28[-<size>]       w:kern w:val="28" (14pt and up): 11pt below it, 16pt above
  kern-none-c14, kern-2-c14   compatibility mode 14 (the others 15)
  liga-standard          w14:ligatures w:val="standard" on every run (11pt)
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
     'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"')
FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Cambria']
WORDS = ['AVAVAVAVAV', 'WAWAWAWAWA', 'ToToToToTo', 'LTLTLTLTLT', 'YoYoYoYoYo', 'Reference', 'office', 'fifty', 'flow']


def para(font, word, size, rpr_extra):
    fonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    return (f'<w:p><w:r><w:rPr>{fonts}<w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/>{rpr_extra}</w:rPr>'
            f'<w:t>{word}</w:t></w:r></w:p>')


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>'''


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def write(path, size, rpr_extra, mode):
    body = ''.join(para(font, word, size, rpr_extra) for font in FONTS for word in WORDS)
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings(mode))


FIXTURES = {
    'kern-none': (11, '', 15),
    'kern-none-16': (16, '', 15),
    'kern-2': (11, '<w:kern w:val="2"/>', 15),
    'kern-2-16': (16, '<w:kern w:val="2"/>', 15),
    'kern-28': (11, '<w:kern w:val="28"/>', 15),
    'kern-28-16': (16, '<w:kern w:val="28"/>', 15),
    'kern-none-c14': (11, '', 14),
    'kern-2-c14': (11, '<w:kern w:val="2"/>', 14),
    'liga-standard': (11, '<w14:ligatures w14:val="standard"/>', 15),
}


def widths(folder):
    word = os.path.join(folder, 'word')
    rows = []
    for name, (size, _, _) in sorted(FIXTURES.items()):
        pdf = os.path.join(word, name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        found = [(m.group(3), float(m.group(2)) - float(m.group(1))) for m in
                 re.finditer(r'xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="[\d.]+">(.*?)</word>', out)]
        expected = [(font, w) for font in FONTS for w in WORDS]
        if [w for w, _ in found] != [w for _, w in expected]:
            print(name, 'words not as written:', [w for w, _ in found][:12], file=sys.stderr)
            continue
        for (font, w), (_, width) in zip(expected, found):
            rows.append(f'{name}\t{font}\t{size}\t{w}\t{width:.3f}')
    with open(os.path.join(word, 'widths.tsv'), 'w') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} widths')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--widths':
        widths(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/kerning')
    os.makedirs(out_dir, exist_ok=True)
    for name, (size, extra, mode) in FIXTURES.items():
        write(os.path.join(out_dir, name + '.docx'), size, extra, mode)
    print(f'{len(FIXTURES)} fixtures in {out_dir}')
