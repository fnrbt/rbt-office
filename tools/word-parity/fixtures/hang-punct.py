#!/usr/bin/env python3
"""Fixtures for punctuation that Word lets run past the end of a line (docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/hang-punct.py [out dir]   (default tests/word/fixtures/hang-punct)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/hang-punct \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/hang-punct --refresh-word
    python3 tools/word-parity/fixtures/long-word.py --lines tests/word/fixtures/hang-punct

(the lines of Word's PDFs in <dir>/word/lines.tsv). tests/word/hang-punct.fsx checks ours against them.

docx4j-getting-started's "... do so directly (see below)." is one line in Word, its text 0.375pt
wider than the line, under its East Asian settings (w:useFELayout, w:lang w:eastAsia="ko-KR",
w:characterSpacingControl doNotCompress); alone it wraps (line-fit's facade-c14). Every fixture
here is one A4 page with 1in margins (a 451.35pt line), Calibri 11pt (2048 units an em, no kerning),
single spacing, nothing after. Every paragraph is a probe: a label, words of 'e's and a last word of
letters chosen so that the text before its probe characters P is exactly 451.35 - s points wide
(to Calibri's font units: within 0.01pt), then P. P fits by its own advance or the last word goes
to the next line - unless P may hang past the margin. The probes (the paragraph's label first):

  c<nn>    one character P, s = 0.40 (0.4pt to spare before P - more than ours' 0.25px fit safety -
           and P over the line by its width less 0.4): . , ) ] } ; : ! ? ” ’ " ' % - » … and
           the controls ( x a
  m<nn>    P, s = 0.40, then " eeee eeee" (P inside the paragraph, the next word on the next line)
  w<nn>    full-width P (。 、 ， ． ） 」 ！, Malgun Gothic, 11pt each) after s = 0.40
  d<nn>    two or three: ). .. .) .” ... ,, ?. )x
  t<nn>    no P: the text itself over the line by 0.05, 0.2, 0.4, 1.0pt (a tolerance, not hanging?)
  i<nn>    c01 (".") with a right indent of 1in: the line ends at the indent, not the margin
  k<nn>    c01 and c03 with a Hangul word on the paragraph's second line (East Asian text in the paragraph)
  x<nn>    the rest of Word's kinsoku lists (Korean, Japanese, Chinese) and a few more (EXTRA), in a
           Malgun Gothic run
  e<nn>    "." then a space at the paragraph's end (e01), then a line break (e02); over the line by
           1.78, 0.78, 0.08pt (e03 - e05); mid-paragraph (e06)
  r<nn>    "." / ")" in a run of their own, at the end (r01, r02) or before " eeee" (r03, r04)
  s<nn>    "?" / "." at 36pt (16.7 / 8.9pt past the line)
  f<nn>    (round 3) "." in a field's result: the probe's text and it (f01), with w:lang w:val="en-US"
           (f02), the "." alone, bold, en-US (f03); b<nn> the "." in a run of its own, bold and en-US,
           en-US, bold (b01 - b03) - docx4j-headers-footers' "Error! Reference source not found."

Round 3 puts the probes where a line ends elsewhere than at the margin - a table cell, a column, a
frame, a header (WHERE, below).

The documents vary the settings that could switch it on, one at a time (compatibility mode 14
unless said):

  none, none-c15            no East Asian setting
  fe3, fe3-c15              docx4j-getting-started's three: useFELayout, eastAsia ko-KR (docDefaults), doNotCompress
  fe3-c12, fe3-nocompat     fe3 in compatibility mode 12, with no compatibilityMode
  usefe, lang, dnc          one of them only;  usefe-lang, usefe-dnc, lang-dnc  two of them
  lang-ja, lang-zh, lang-zhtw, lang-en   fe3 with w:eastAsia ja-JP, zh-CN, zh-TW, en-US
  fe3-op0, fe3-op0-style    fe3 with w:overflowPunct w:val="0" in docDefaults' pPr / on the Normal style
  op1                       none with w:overflowPunct (on) in docDefaults' pPr
  fe3-cp, fe3-cpk           fe3 with compressPunctuation / compressPunctuationAndJapaneseKana
  lang-ja-cp                lang-ja with compressPunctuation
  runlang                   usefe and doNotCompress, eastAsia ko-KR on every run instead of docDefaults
  lastrun, notlast          the same, ko-KR only on the r probes' punctuation run / only on the run before it
(Compatibility mode 13 is not one: Word does not open such a document.)
No dependencies: the package is written as XML; Calibri's advances are below.
"""
import os
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
# Calibri's advance widths, font units of 2048 an em (calibri.ttf of Windows 10).
ADV = {'a': 981, 'b': 1076, 'c': 866, 'd': 1076, 'e': 1019, 'f': 625, 'g': 964, 'h': 1076, 'i': 470, 'j': 490,
       'k': 931, 'l': 470, 'm': 1636, 'n': 1076, 'o': 1080, 'p': 1076, 'q': 1076, 'r': 714, 's': 801, 't': 686,
       'u': 1076, 'v': 925, 'w': 1464, 'x': 887, 'y': 927, 'z': 809, '0': 1038, '1': 1038, '2': 1038, '3': 1038,
       '4': 1038, '5': 1038, '6': 1038, '7': 1038, '8': 1038, '9': 1038, ' ': 463, '.': 517, ',': 511, ')': 621,
       ']': 628, '}': 644, ';': 548, ':': 548, '!': 667, '?': 949, '\u201d': 857, '\u2019': 511, '"': 821, "'": 452,
       '%': 1464, '-': 627, '(': 621, '\u00bb': 1049, '\u2026': 1414}
LINE = 9027  # twips: 11907 less two 1440 margins
PT = 11 / 2048


def width(t):
    return sum(ADV[c] for c in t) * PT


def last_word(target_units):
    """Letters whose advances sum to target_units or a little under: two at most of each of
    j f t r s z c, the rest 'e's, the combination nearest to it."""
    import itertools
    fine = 'jftrszc'
    best = (target_units + 1, '')
    for counts in itertools.product(range(3), repeat=len(fine)):
        u = sum(n * ADV[c] for n, c in zip(counts, fine))
        if u > target_units:
            continue
        e, rest = divmod(target_units - u, ADV['e'])
        if rest < best[0]:
            best = (rest, 'e' * e + ''.join(c * n for c, n in zip(fine, counts)))
    return best[1]


def probe_text(label, s, line_pt=LINE / 20):
    """label + e-words + a last word: line_pt - s points wide (to the nearest unit under)."""
    total = round((line_pt - s) / PT)
    pre = label + ' '
    while total - sum(ADV[c] for c in pre + 'eeeeeeee ') > 8000:
        pre += 'eeeeeeee '
    return pre + last_word(total - sum(ADV[c] for c in pre))


def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def run(t, rpr=''):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>'


MALGUN = '<w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic" w:cs="Malgun Gothic"/>'
CHARS = ['.', ',', ')', ']', '}', ';', ':', '!', '?', '\u201d', '\u2019', '"', "'", '%', '-', '\u00bb', '\u2026', '(', 'x', 'a']
WIDE = ['\u3002', '\u3001', '\uff0c', '\uff0e', '\uff09', '\u300d', '\uff01']
MULTI = [').', '..', '.)', '.\u201d', '...', ',,', '?.', ')x']
OVER = [0.05, 0.2, 0.4, 1.0]
# Round 2: the rest of Word's default kinsoku lists (characters that may not start a line) for Korean,
# Japanese and Chinese, and a few that are in none of them, every one in a Malgun Gothic run.
EXTRA = ['\u00a2', '\u00b0', '\u2032', '\u2033', '\u2103', '\u3009', '\u300b', '\u300f', '\u3011', '\u3015',
         '\uff05', '\uff1a', '\uff1b', '\uff1f', '\uff3d', '\uff5d', '\uffe0', '\u00b7', '\u2030', '\u203a',
         '\u3003', '\u3017', '\u301e', '\uff5e', '\uff02', '\uff07', '\u3005', '\u309d', '\u30fb', '\u30fc',
         '\u3041', '\u30a1', '\uff61', '\uff63', '\uff64', '\u2010', '\u2016', '\u2015', '\u02c7', '\u00a8',
         ')', '.', '!', '\u201d', '\u2019', '\u2026', '\u300c', '\uff08']


def probes(rpr, rpr_t=None, rpr_p=None, line_pt=LINE / 20, ppr0='', fields=False):
    """(kind, paragraph xml) for every probe, on lines line_pt wide; ppr0 in every paragraph's
    w:pPr (a frame's w:framePr); fields: the f and b probes too."""
    out = []
    p = lambda body, ppr='': f'<w:p><w:pPr>{ppr0}{ppr}</w:pPr>{body}</w:p>'
    pt = lambda label, s: probe_text(label, s, line_pt)
    for n, c in enumerate(CHARS, 1):
        out.append(('c', p(run(pt(f'c{n:02}', 0.40) + c, rpr))))
    for n, c in enumerate(['.', ')', ','], 1):
        out.append(('m', p(run(pt(f'm{n:02}', 0.40) + c + ' eeee eeee', rpr))))
    for n, c in enumerate(WIDE, 1):
        out.append(('w', p(run(pt(f'w{n:02}', 0.40), rpr) + run(c, MALGUN + rpr))))
    for n, c in enumerate(MULTI, 1):
        out.append(('d', p(run(pt(f'd{n:02}', 0.40) + c, rpr))))
    for n, o in enumerate(OVER, 1):
        out.append(('t', p(run(pt(f't{n:02}', -o), rpr))))
    out.append(('i', p(run(probe_text('i01', 0.40, line_pt - 72) + '.', rpr), '<w:ind w:right="1440"/>')))
    for n, c in enumerate(['.', ')'], 1):
        out.append(('k', p(run(pt(f'k{n:02}', 0.40) + c + ' eeee ', rpr) + run('\uac00\ub098\ub2e4', MALGUN + rpr))))
    # round 2
    for n, c in enumerate(EXTRA, 1):
        out.append(('x', p(run(pt(f'x{n:02}', 0.40), rpr) + run(c, MALGUN + rpr))))
    out.append(('e', p(run(pt('e01', 0.40) + '. ', rpr))))                       # a space after it
    out.append(('e', p(run(pt('e02', 0.40) + '.', rpr) + '<w:r><w:br/></w:r>' + run('eeee', rpr))))  # a line break
    for n, sl in enumerate([1.0, 2.0, 2.7], 3):                                 # '.' over by 1.78, 0.78, 0.08
        out.append(('e', p(run(pt(f'e{n:02}', sl) + '.', rpr))))
    out.append(('e', p(run(pt('e06', 0.40) + '. eeee', rpr))))                  # mid-paragraph, as m01
    rt = rpr if rpr_t is None else rpr_t
    rp = rpr if rpr_p is None else rpr_p
    for n, c in enumerate(['.', ')'], 1):                                        # the probe in a run of its own
        out.append(('r', p(run(pt(f'r{n:02}', 0.40), rt) + run(c, rp))))
        out.append(('r', p(run(pt(f'r{n + 2:02}', 0.40), rt) + run(c + ' eeee', rp))))
    big = '<w:sz w:val="72"/><w:szCs w:val="72"/>'                              # '?' at 36pt: 16.7pt over
    out.append(('s', p(run(pt('s01', 0.40), rpr) + run('?', big + rpr))))
    out.append(('s', p(run(pt('s02', 0.40), rpr) + run('.', big + rpr))))
    if fields:
        # round 3: docx4j-headers-footers' "Error! Reference source not found." - a REF field's result,
        # bold, w:lang w:val="en-US" (the East Asian language still ko-KR from docDefaults).
        en = '<w:lang w:val="en-US"/>'
        bold = '<w:b/><w:bCs/>'
        out.append(('f', p(complex_field(' QUOTE "x" ', run(pt('f01', 0.40) + '.', rpr)))))
        out.append(('f', p(complex_field(' QUOTE "x" ', run(pt('f02', 0.40) + '.', en + rpr)))))
        out.append(('f', p(run(pt('f03', 0.40), rpr) + complex_field(' QUOTE "x" ', run('.', bold + en + rpr)))))
        out.append(('f', p(complex_field(' REF nowhere \\h ', run(pt('f04', 0.40) + '.', rpr)))))
        out.append(('b', p(run(pt('b01', 0.40), rpr) + run('.', bold + en + rpr))))
        out.append(('b', p(run(pt('b02', 0.40), rpr) + run('.', en + rpr))))
        out.append(('b', p(run(pt('b03', 0.40), rpr) + run('.', bold + rpr))))
    return out


def complex_field(code, result_runs):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve">{esc(code)}</w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r>{result_runs}<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def styles(lang=None, op=None, op_style=None):
    lang_x = f'<w:lang w:val="en-AU" w:eastAsia="{lang}" w:bidi="ar-SA"/>' if lang else ''
    op_x = f'<w:overflowPunct w:val="{op}"/>' if op is not None else ''
    st_x = f'<w:pPr><w:overflowPunct w:val="{op_style}"/></w:pPr>' if op_style is not None else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
            f'<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>{lang_x}</w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr>{op_x}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{st_x}</w:style></w:styles>')


def settings(mode, usefe=False, csc=None):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            + (f'<w:characterSpacingControl w:val="{csc}"/>' if csc else '') + '<w:compat>'
            + ('<w:useFELayout/>' if usefe else '')
            + (f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>' if mode else '')
            + '</w:compat></w:settings>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
SECT = ('<w:sectPr><w:pgSz w:w="11907" w:h="16839" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')

DNC = 'doNotCompress'
# name: (compatibility mode, useFELayout, characterSpacingControl, docDefaults' eastAsia, overflowPunct in
# docDefaults, on Normal, eastAsia on the runs)
DOCS = {
    'none': (14, False, None, None, None, None, None),
    'none-c15': (15, False, None, None, None, None, None),
    'fe3': (14, True, DNC, 'ko-KR', None, None, None),
    'fe3-c15': (15, True, DNC, 'ko-KR', None, None, None),
    'usefe': (14, True, None, None, None, None, None),
    'lang': (14, False, None, 'ko-KR', None, None, None),
    'dnc': (14, False, DNC, None, None, None, None),
    'usefe-lang': (14, True, None, 'ko-KR', None, None, None),
    'usefe-dnc': (14, True, DNC, None, None, None, None),
    'lang-dnc': (14, False, DNC, 'ko-KR', None, None, None),
    'lang-ja': (14, True, DNC, 'ja-JP', None, None, None),
    'lang-zh': (14, True, DNC, 'zh-CN', None, None, None),
    'lang-en': (14, True, DNC, 'en-US', None, None, None),
    'fe3-op0': (14, True, DNC, 'ko-KR', 0, None, None),
    'fe3-op0-style': (14, True, DNC, 'ko-KR', None, 0, None),
    'op1': (14, False, None, None, 1, None, None),
    'fe3-cp': (14, True, 'compressPunctuation', 'ko-KR', None, None, None),
    'runlang': (14, True, DNC, None, None, None, 'ko-KR'),
    # round 2
    'lang-zhtw': (14, True, DNC, 'zh-TW', None, None, None),
    'fe3-c12': (12, True, DNC, 'ko-KR', None, None, None),
    'fe3-nocompat': (None, True, DNC, 'ko-KR', None, None, None),
    'fe3-cpk': (14, True, 'compressPunctuationAndJapaneseKana', 'ko-KR', None, None, None),
    'lang-ja-cp': (14, True, 'compressPunctuation', 'ja-JP', None, None, None),
    'lastrun': (14, True, DNC, None, None, None, None, 'last'),
    'notlast': (14, True, DNC, None, None, None, None, 'notlast'),
}


def write(path, mode, usefe, csc, lang, op, op_style, run_lang, split=None):
    rpr = f'<w:lang w:val="en-AU" w:eastAsia="{run_lang}"/>' if run_lang else ''
    ko = '<w:lang w:val="en-AU" w:eastAsia="ko-KR"/>'
    body = ''.join(x for _, x in probes(rpr, *({'last': ('', ko), 'notlast': (ko, '')}[split] if split else (None, None))))
    package(path, body, styles(lang, op, op_style), settings(mode, usefe, csc))


R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
HDR_TYPE = '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
HDR_REL = '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'


def package(path, body, styles_xml, settings_xml, sect=SECT, header=None):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>{body}{sect}</w:body></w:document>'
    types = TYPES if header is None else TYPES.replace('</Types>', HDR_TYPE + '</Types>')
    rels = DOC_RELS if header is None else DOC_RELS.replace('</Relationships>', HDR_REL + '</Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', types)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', rels)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles_xml)
        z.writestr('word/settings.xml', settings_xml)
        if header is not None:
            z.writestr('word/header1.xml', f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{header}</w:hdr>')


def table(paragraphs, tcw, mar, layout='fixed', tblw=None):
    """One cell of tcw twips, margins of mar twips left and right (none above and below), no borders."""
    w = tblw or f'<w:tblW w:w="{tcw}" w:type="dxa"/>'
    lay = '<w:tblLayout w:type="fixed"/>' if layout == 'fixed' else ''
    return (f'<w:tbl><w:tblPr>{w}{lay}<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
            f'<w:left w:w="{mar}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="{mar}" w:type="dxa"/>'
            f'</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="{tcw}"/></w:tblGrid><w:tr><w:tc><w:tcPr>'
            f'<w:tcW w:w="{tcw}" w:type="dxa"/></w:tcPr>{paragraphs}</w:tc></w:tr></w:tbl><w:p/>')


def pick(ps, kinds, n=None):
    xs = [x for k, x in ps if k in kinds]
    return ''.join(xs[:n] if n else xs)


# Round 3: where the line ends - a table cell's text, a column, a frame, a header (docx4j-headers-footers'
# narrow column: "not" / "found." in Word, the period not hung past the cell's text). Every one under fe3's
# settings (compatibility mode 14, useFELayout, ko-KR, doNotCompress), with the f and b probes (fields,
# bold, w:lang w:val="en-US" on the punctuation's run):
#   fe3-fld        the body, as fe3: the f and b probes only
#   cell           a fixed one-cell table, 300pt of text (cell margins 0)
#   cell-mar       the same with 0.075in (108 twips) cell margins, the cell 6216 twips: 300pt of text
#   cell-edge      a fixed one-cell table of the page's text width, margins 0: the cell's text ends at the margin
#   cell-auto      an autofit 100% table, 108-twip margins (in compatibility mode 14 out by them: its text the page's)
#   cell-c15       cell in compatibility mode 15 (no hanging anywhere: the control)
#   cols2          two columns of 207.5pt (727 twips between), the probes in the first of each page
#   frame          paragraphs in one frame 300pt wide (w:framePr w:w="6000"), at the margin
#   header         the c probes 1-10 in the header (the page's text width)
WHERE = ['fe3-fld', 'cell', 'cell-mar', 'cell-edge', 'cell-auto', 'cell-c15', 'cols2', 'frame', 'header']
MAIN = 'cmwdtiekrsfb'


def write_where(path, name):
    mode = 15 if name == 'cell-c15' else 14
    st, se = styles('ko-KR'), settings(mode, True, DNC)
    fe = lambda line_pt=LINE / 20, ppr0='': probes('', line_pt=line_pt, ppr0=ppr0, fields=True)
    if name == 'fe3-fld':
        package(path, pick(fe(), 'fb'), st, se)
    elif name in ('cell', 'cell-c15'):
        package(path, table(pick(fe(300), MAIN), 6000, 0), st, se)
    elif name == 'cell-mar':
        package(path, table(pick(fe(300), MAIN), 6216, 108), st, se)
    elif name == 'cell-edge':
        package(path, table(pick(fe(), MAIN), LINE, 0), st, se)
    elif name == 'cell-auto':
        package(path, table(pick(fe(), MAIN), LINE, 108, 'auto', '<w:tblW w:w="5000" w:type="pct"/>'), st, se)
    elif name == 'cols2':
        # (12 probes in the first column of each page, the second column kept empty - two column
        # breaks - so that no line of Word's PDF has words of both.)
        xs = [x for k, x in fe(207.5) if k in MAIN]
        brk = '<w:p><w:r><w:br w:type="column"/><w:br w:type="column"/></w:r></w:p>'
        sect = SECT.replace('</w:sectPr>', '<w:cols w:num="2" w:space="727"/></w:sectPr>')
        package(path, brk.join(''.join(xs[i:i + 12]) for i in range(0, len(xs), 12)), st, se, sect)
    elif name == 'frame':
        fr = '<w:framePr w:w="6000" w:wrap="around" w:hAnchor="margin" w:vAnchor="margin" w:x="0" w:y="0"/>'
        package(path, pick(fe(300, fr), 'cmwdt', 40) + '<w:p/>', st, se)
    elif name == 'header':
        sect = SECT.replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>')
        package(path, '<w:p/>', st, se, sect, pick(fe(), 'c', 10))


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/hang-punct')
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    for name, spec in DOCS.items():
        write(os.path.join(out_dir, name + '.docx'), *spec)
    for name in WHERE:
        write_where(os.path.join(out_dir, name + '.docx'), name)
    print(f'{len(DOCS) + len(WHERE)} fixtures in {out_dir}')
