#!/usr/bin/env python3
"""Fixtures for where Word breaks a line inside a word: after a hyphen, a dash, a slash ... (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/hyphen-break.py [out dir]   (default tests/word/fixtures/hyphen-break)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/hyphen-break \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/hyphen-break --refresh-word
    python3 tools/word-parity/fixtures/long-word.py --lines tests/word/fixtures/hyphen-break
    python3 tools/word-parity/fixtures/hyphen-break.py --table tests/word/fixtures/hyphen-break

(the lines of Word's PDFs in <dir>/word/lines.tsv; --table prints, per probe and document, where
Word broke). tests/word/hyphen-break.fsx checks ours against them.

Made to ask whether docx4j-getting-started's "http://www.ecma-international.org/news/..." may break
after "ecma-" (it does, in Word's PDF of it too: page 12 starts "international.org/news/"). Every fixture is Courier New 10pt (6.0pt a character, no kerning),
exactly 20pt a line, on a Letter page with 1in margins and a right indent leaving 363pt: a line holds
60 characters and not 61. Every paragraph is a probe: a label and 'a's (40 characters), a space, a
token of 29 or so characters with the probed character P at its 13th (the line's 54th) and " end".
The token does not fit after the prefix but fits a line of its own, so the first line ends
  - after P when Word may break after it,
  - before P when it may break before it,
  - with the 'a's when neither (the token goes to the next line whole).
The probes:
  a<nn>   P between letters: - U+2010 U+2011 w:noBreakHyphen w:softHyphen U+00AD U+2012 U+2013 U+2014
          U+2015 U+2212 / \\ . , : ; _ ? ! & = + % ( ) ~ | * # @ U+00B7
  d<nn>   P between digits: - U+2013 U+2014 / . , : w:softHyphen w:noBreakHyphen
  x<nn>   "-" between a letter and a digit, a digit and a letter
  u<nn>   "-" between letters in a token with "." "/" ":" "@" "_" "," "\\" beside it or further off,
          two hyphens, web addresses ("http://www.ecma-international.org/news/...")
  b<nn>   P the line's 61st character, the 60 before it filling the line: - U+2010 U+2013 U+2014
          w:softHyphen / ( \\ U+2011 (a break before P, or must P fit?); b10 "-" the 60th
  l<nn>   a hyphen, em or en dash at the token's start (before letters, before digits), after one
          letter, "--", "-)", "(" or '"' before the word
  s<nn>   a soft hyphen not broken at (s01: what width?), two in a word (s02)
The documents (one setting each):
  base     compatibility mode 15, en-US
  c14      compatibility mode 14
  fe3      docx4j-getting-started's East Asian settings: mode 14, w:useFELayout, w:lang en-AU / eastAsia ko-KR,
           doNotCompress
  link     each token in a w:hyperlink (an anchor)
  extlink  each token in a w:hyperlink to an external address, in the Hyperlink character style
  fld      each token as a HYPERLINK field's result
No dependencies: the package is written as XML.
"""
import os
import re
import sys
import zipfile

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
IND = '<w:ind w:right="2100"/>'  # 468 - 105 = 363pt: 60 characters of 6pt
PREFIX = 40

L, R = 'b' * 12, 'c' * 16
LD, RD = '1' * 12, '2' * 16
NBH, SHY = '<NBH>', '<SHY>'  # w:noBreakHyphen, w:softHyphen

LETTER_CHARS = ['-', '\u2010', '\u2011', NBH, SHY, '\u00ad', '\u2012', '\u2013', '\u2014', '\u2015', '\u2212',
                '/', '\\', '.', ',', ':', ';', '_', '?', '!', '&', '=', '+', '%', '(', ')', '~', '|', '*', '#', '@',
                '\u00b7']
DIGIT_CHARS = ['-', '\u2013', '\u2014', '/', '.', ',', ':', SHY, NBH]
EDGE_CHARS = ['-', '\u2010', '\u2013', '\u2014', SHY, '/', '(', '\\', '\u2011']
LEADING = [
    '-' + 'c' * 28,                    # l01 a hyphen at the token's start
    '\u2014' + 'c' * 28,               # l02 an em dash at its start
    '\u2013' + 'c' * 28,               # l03 an en dash at its start
    'b-' + 'c' * 27,                   # l04 one letter before the hyphen
    'b' * 12 + '--' + 'c' * 15,        # l05 two hyphens (the 54th and 55th)
    '-' + '2' * 28,                    # l06 a minus sign before digits
    'b' * 12 + '-)' + 'c' * 15,        # l07 a hyphen before ")"
    '(' + 'b' * 11 + '-' + 'c' * 16,   # l08 "(" before
    '"' + 'b' * 11 + '-' + 'c' * 16,   # l09 a quote before
]
# A soft hyphen not broken at (s01: its width?), two in a word (s02: the later one that fits).
SOFT = {
    's01': 'bbbb' + SHY + 'cccc',
    's02': 'b' * 6 + SHY + 'b' * 6 + SHY + 'c' * 16,
}
CONTEXTS = [
    'bbbbbbb.bbbb-cccccccccccccccc',   # u01 "." before
    'bbbbbbbbbbbb-cccccc.ccccccccc',   # u02 "." after
    'bbbbbbb/bbbb-cccccccccccccccc',   # u03 "/" before
    'bbbbbbbbbbbb-cccccc/ccccccccc',   # u04 "/" after
    'bbbbbbb:bbbb-cccccccccccccccc',   # u05 ":" before
    'bbbbbbbbbbbb-cccccc:ccccccccc',   # u06 ":" after
    'bbbbbbb@bbbb-cccccccccccccccc',   # u07 "@" before
    'bbbbbbb_bbbb-cccccccccccccccc',   # u08 "_" before
    'bbbbbbb,bbbb-cccccccccccccccc',   # u09 "," before
    'http://w.ecm-cccccccccccccccc',   # u10 a scheme
    'www.ecmaabcd-international.org',  # u11 "www." and ".org"
    'ttp//wwwecma-cccccccccccccccc',   # u12 "//" only
    'http://www.ecma-international.org/news/',  # u13 getting-started's (hyphen the 16th)
    'bbbbbbbbbbbb-3333333333.htm',     # u14 "Ecma-376.htm"
    'C:\\dir\\bbbbb-cccccccccccccccc',  # u15 a path
    'b.bbbbbbbbbb-cccccccccccccccc',   # u16 "." far before
    'bbbbbbbbbbbb-ccccccccccccccc.c',  # u17 "." far after
    'bbbbbbb-bbbb-cccccccccccccccc',   # u18 two hyphens
    'bbbbbbbbbbb.-cccccccccccccccc',   # u19 ".-"
    'bbbbbbbbbbbb-.ccccccccccccccc',   # u20 "-."
    'bbbbbbbbbbbb-/ccccccccccccccc',   # u21 "-/"
    'bbbbbbbbbbb/-cccccccccccccccc',   # u22 "/-"
]


def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def r(text, rpr=''):
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def runs(text, rpr=''):
    """`text` as runs, NBH / SHY as w:noBreakHyphen / w:softHyphen."""
    out = ''
    for part in re.split(f'({NBH}|{SHY})', text):
        if part == NBH:
            out += f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:noBreakHyphen/></w:r>'
        elif part == SHY:
            out += f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:softHyphen/></w:r>'
        elif part:
            out += r(part, rpr)
    return out


def probes():
    """label -> token (NBH / SHY markers in it)."""
    out = {}
    for i, c in enumerate(LETTER_CHARS, 1):
        out[f'a{i:02d}'] = L + c + R
    for i, c in enumerate(DIGIT_CHARS, 1):
        out[f'd{i:02d}'] = LD + c + RD
    out['x01'] = L + '-' + RD
    out['x02'] = LD + '-' + R
    for i, t in enumerate(CONTEXTS, 1):
        out[f'u{i:02d}'] = t
    # P the line's 61st character (what is before it fills the line): a break before P, or none.
    for i, c in enumerate(EDGE_CHARS, 1):
        out[f'b{i:02d}'] = 'b' * 19 + c + 'c' * 10
    out['b10'] = 'b' * 18 + '-' + 'c' * 11  # the hyphen the 60th: it fits
    for i, t in enumerate(LEADING, 1):
        out[f'l{i:02d}'] = t
    out.update(SOFT)
    return out


def wrap(kind, token):
    if kind == 'link':
        return f'<w:hyperlink w:anchor="missing">{runs(token)}</w:hyperlink>'
    if kind == 'extlink':
        return f'<w:hyperlink r:id="rId9" w:history="1">{runs(token, "<w:rStyle w:val=\"Hyperlink\"/>")}</w:hyperlink>'
    if kind == 'fld':
        return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> HYPERLINK "http://www.example.com/" </w:instrText></w:r>'
                f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>{runs(token)}<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    return runs(token)


def body(kind):
    out = ''
    for label, token in probes().items():
        prefix = label + 'a' * (PREFIX - len(label)) + ' '
        out += f'<w:p><w:pPr>{IND}</w:pPr>{r(prefix)}{wrap(kind, token)}{r(" end")}</w:p>'
    return out


def styles(lang):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
            '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/>'
            f'<w:sz w:val="20"/><w:szCs w:val="20"/>{lang}</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr></w:style>'
            '</w:styles>')


def settings(mode, usefe, csc):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            + (f'<w:characterSpacingControl w:val="{csc}"/>' if csc else '') + '<w:compat>'
            + ('<w:useFELayout/>' if usefe else '')
            + f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            + '</w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.example.com/" TargetMode="External"/></Relationships>'''
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')

EN = '<w:lang w:val="en-US"/>'
FE = '<w:lang w:val="en-AU" w:eastAsia="ko-KR" w:bidi="ar-SA"/>'
# name: (kind, compatibility mode, useFELayout, characterSpacingControl, docDefaults' w:lang)
DOCS = {
    'base': ('plain', 15, False, None, EN),
    'c14': ('plain', 14, False, None, EN),
    'fe3': ('plain', 14, True, 'doNotCompress', FE),
    'link': ('link', 15, False, None, EN),
    'extlink': ('extlink', 15, False, None, EN),
    'fld': ('fld', 15, False, None, EN),
}


def write(path, kind, mode, usefe, csc, lang):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body(kind)}{SECT}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles(lang))
        z.writestr('word/settings.xml', settings(mode, usefe, csc))


def table(folder):
    """Where Word broke each probe, from <folder>/word/lines.tsv: the first line's length past the prefix
    (0: the token went to the next line; 12: before P; 13: after P)."""
    lines = {}
    for l in open(os.path.join(folder, 'word/lines.tsv'), encoding='utf-8').read().splitlines():
        c = l.split('\t')
        lines.setdefault(c[0], []).append(c[3])
    docs = [d for d in DOCS if d in lines]
    print('probe\ttoken\t' + '\t'.join(docs))
    for label, token in probes().items():
        row = []
        for d in docs:
            ls = lines[d]
            i = next((k for k, t in enumerate(ls) if t.startswith(label + 'a')), None)
            if i is None:
                row.append('?')
                continue
            first = ls[i].split(' ', 1)
            row.append(str(len(first[1])) if len(first) > 1 else '0')
        shown = token.replace(NBH, '[nbh]').replace(SHY, '[shy]').replace('\u00ad', '[ad]')
        print(f'{label}\t{shown}\t' + '\t'.join(row))


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--table':
        table(sys.argv[2])
        sys.exit(0)
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/hyphen-break')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    for name, spec in DOCS.items():
        write(os.path.join(out_dir, name + '.docx'), *spec)
    print(f'{len(DOCS)} fixtures in {out_dir}')
