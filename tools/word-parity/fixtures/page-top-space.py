#!/usr/bin/env python3
"""Fixtures for a paragraph's space before at the top of a page or column (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/page-top-space.py [out dir]   (default tests/word/fixtures/page-top-space)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/page-top-space \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/page-top-space --refresh-word
    python3 tools/word-parity/fixtures/page-top-space.py --tops tests/word/fixtures/page-top-space

--tops reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/tops.tsv: for
every probe word ("P-..."), the fixture, its page, the left and the top of the word in points.
tests/word/page-top-space.fsx checks ours against them.

docx4j-headers-footers' "pgMar (Page Margins)" (Heading 3, 8pt before) follows a paragraph that is
only a page break (<w:br w:type="page"/>); Word sets it at the top margin, its space before gone, ours
8pt below. Every fixture here is A4 with 1in margins, Calibri 11pt, single spacing, nothing before or
after but on the probes: a probe is a paragraph "P-<name> ..." with 12pt before, and what puts it at
the top of a page (or a column) is what varies, one per page:

  first        the document's first paragraph
  brkpara      after a paragraph that is only a page break (as docx4j-headers-footers)
  brkpara-kn   the same, the probe keepNext and keepLines (as its Heading 3)
  brkend       after a paragraph whose text ends in a page break
  brkin        the page break at the start of the probe's own paragraph
  brkmid       the page break inside the probe's paragraph (12pt before), after its first words
  pbb          the probe has w:pageBreakBefore
  flow         pushed to the page by the 51 lines above it (the page holds 51.97 lines; the first pageBreakBefore)
  brkpara-nosp the paragraph-only page break, the probe with no space before (the control: the top)

in compatibility mode 14 (pages-c14) and 15 (pages-c15), one section; the same in the second section
of a document of two (sect-c14, sect-c15: such a document is laid out section by section), its first
probe "s-sect" after a next-page section break; and in a section of two columns, after a
column break in a paragraph of its own (cols-brkpara-c14, -c15) and pushed into the second column by
the lines of the first (cols-flow-c14, -c15). Also: pages-c14-sup, pages-c15-sup with
w:suppressSpBfAfterPgBrk (Word's "Suppress Space Before after a hard page break").
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'''
TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
PG = ('<w:pgSz w:w="11907" w:h="16839" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
      'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>')
COLS = '<w:cols w:num="2" w:space="720"/>'


def settings(mode, sup=False):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            + ('<w:suppressSpBfAfterPgBrk/>' if sup else '') +
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


def p(text, ppr='', pre=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr>{pre}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


BEFORE = '<w:spacing w:before="240"/>'
BRK = '<w:r><w:br w:type="page"/></w:r>'


def probe(name, ppr=BEFORE, pre=''):
    return p(f'P-{name} probe', ppr, pre) + p(f'after {name}')


def pages(x=''):
    """The probes, one a page; x before their names (and the first: 'sect' after a section break)."""
    out = probe(x + 'first' if not x else x + 'sect')
    out += f'<w:p>{BRK}</w:p>' + probe(x + 'brkpara')
    out += f'<w:p>{BRK}</w:p>' + probe(x + 'brkpara-kn', '<w:keepNext/><w:keepLines/>' + BEFORE)
    out += p('body end').replace('</w:t></w:r></w:p>', f'</w:t></w:r>{BRK}</w:p>') + probe(x + 'brkend')
    out += probe(x + 'brkin', pre=BRK)
    out += (f'<w:p><w:pPr>{BEFORE}</w:pPr><w:r><w:t xml:space="preserve">mid start</w:t></w:r>{BRK}'
            f'<w:r><w:t xml:space="preserve">P-{x}brkmid probe</w:t></w:r></w:p>' + p(f'after {x}brkmid'))
    out += probe(x + 'pbb', '<w:pageBreakBefore/>' + BEFORE)
    out += ''.join(p(f'fill {n}', '<w:pageBreakBefore/>' if n == 1 else '') for n in range(1, 52)) + probe(x + 'flow')
    out += f'<w:p>{BRK}</w:p>' + probe(x + 'brkpara-nosp', '')
    return out + f'<w:sectPr>{PG}</w:sectPr>'


def sect():
    return p('section end', f'<w:sectPr>{PG}</w:sectPr>') + pages('s-')


def cols(kind):
    if kind == 'brkpara':
        body = p('column one') + '<w:p><w:r><w:br w:type="column"/></w:r></w:p>' + probe('col-brkpara')
    else:
        body = ''.join(p(f'fill {n}') for n in range(1, 52)) + probe('col-flow')
    return body + f'<w:sectPr>{PG}{COLS}</w:sectPr>'


def docs():
    out = {}
    for mode in (14, 15):
        out[f'pages-c{mode}'] = (pages(), settings(mode))
        out[f'pages-c{mode}-sup'] = (pages(), settings(mode, True))
        out[f'sect-c{mode}'] = (sect(), settings(mode))
        for kind in ('brkpara', 'flow'):
            out[f'cols-{kind}-c{mode}'] = (cols(kind), settings(mode))
    return out


def write(path, body, settings_xml):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/settings.xml', settings_xml)


def tops(folder):
    """<folder>/word/tops.tsv from Word's PDFs: name, page, left, top, word - every P- word."""
    rows = []
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(folder, 'word', name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
        for page, chunk in enumerate(out.split('<page ')[1:], 1):
            for m in re.finditer(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(.*?)</word>', chunk):
                if m.group(3).startswith('P-') or m.group(3) in ('column', 'fill', 'after', 'body', 'section'):
                    rows.append(f'{name}\t{page}\t{float(m.group(1)):.2f}\t{float(m.group(2)):.2f}\t{m.group(3)}')
    with open(os.path.join(folder, 'word', 'tops.tsv'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(rows) + '\n')
    print(f'{len(rows)} words')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--tops':
        tops(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/page-top-space')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    ds = docs()
    for name, (body, st) in ds.items():
        write(os.path.join(out_dir, name + '.docx'), body, st)
    print(f'{len(ds)} fixtures in {out_dir}')
