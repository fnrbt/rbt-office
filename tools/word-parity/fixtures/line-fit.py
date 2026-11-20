#!/usr/bin/env python3
"""Fixtures for how much text Word puts on a line (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/line-fit.py [out dir]   (default tests/word/fixtures/line-fit)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/line-fit \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/line-fit --refresh-word
    python3 tools/word-parity/fixtures/long-word.py --lines tests/word/fixtures/line-fit

(the lines of Word's PDFs in <dir>/word/lines.tsv, as for long-word). Every fixture is one A4 page
with 1in margins (a 451.35pt line, as docx4j-getting-started's last section), single spacing,
nothing after; one paragraph per line to be measured, each a word longer than the line - one
character over and over, so Word breaks it where the line is full and the count of characters on
its first line is what Word took the line to hold - or words of that character with spaces:

  <font>-<size>            one paragraph per character e, a, i, m, W, period, space-separated "ee"
                           words: Calibri 11, Cambria 11, Times New Roman 12, Arial 10, Consolas 10
  <font>-<size>-c15        Calibri 11 in compatibility mode 15 (the others 14)
No dependencies: the package is written as XML.
"""
import os
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CHARS = ['e', 'a', 'i', 'm', 'W', '.']


def para(font, size, text):
    return (f'<w:p><w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{size * 2}"/><w:szCs w:val="{size * 2}"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16839" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')


def settings(mode, fe=False):
    """fe: as docx4j-getting-started - East Asian layout (w:useFELayout), punctuation not compressed."""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            + ('<w:characterSpacingControl w:val="doNotCompress"/>' if fe else '') + '<w:compat>'
            + ('<w:useFELayout/>' if fe else '') +
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


def write(path, font, size, mode):
    body = ''.join(para(font, size, c * 300) for c in CHARS) + para(font, size, ' '.join(['ee'] * 150))
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings(mode))


# docx4j-getting-started's line that Word fits 2.6pt past its margin, and the same changed one thing at a time.
FACADE = "The façade can&apos;t use Microsoft Graph at this time.  If you want to use that, do so directly (see below)."
FACADE_VARIANTS = {
    'facade': FACADE,
    'facade-1space': FACADE.replace('.  If', '. If'),
    'facade-c': FACADE.replace('ç', 'c'),
    'facade-apos': FACADE.replace('can&apos;t', 'cant'),
    'facade-plain': FACADE.replace('.  If', '. If').replace('ç', 'c').replace('can&apos;t', 'cant'),
    'facade-more': FACADE + 'x',
}


def write_facade(path, mode, fe=False):
    body = ''.join(para('Calibri', 11, t) for t in FACADE_VARIANTS.values())
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/settings.xml', settings(mode, fe))
        z.writestr('word/styles.xml', STYLES.replace('<w:sz w:val="22"/>', '<w:sz w:val="22"/><w:lang w:val="en-AU" w:eastAsia="ko-KR" w:bidi="ar-SA"/>') if fe else STYLES)


FIXTURES = {
    'calibri-11': ('Calibri', 11, 14), 'calibri-11-c15': ('Calibri', 11, 15), 'cambria-11': ('Cambria', 11, 14),
    'times-12': ('Times New Roman', 12, 14), 'arial-10': ('Arial', 10, 14), 'consolas-10': ('Consolas', 10, 14),
}

if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/line-fit')
    os.makedirs(out_dir, exist_ok=True)
    for name, (font, size, mode) in FIXTURES.items():
        write(os.path.join(out_dir, name + '.docx'), font, size, mode)
    write_facade(os.path.join(out_dir, 'facade-c14.docx'), 14)
    write_facade(os.path.join(out_dir, 'facade-c15.docx'), 15)
    write_facade(os.path.join(out_dir, 'facade-fe-c14.docx'), 14, fe=True)
    print(f'{len(FIXTURES) + 3} fixtures in {out_dir}')
