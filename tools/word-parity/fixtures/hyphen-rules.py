#!/usr/bin/env python3
"""Fixtures for when Word hyphenates (w:autoHyphenation; docs/layout-parity-flow.md): the zone, the limit
on consecutive hyphens, caps, digits, languages, paragraph and page settings.

    python3 tools/word-parity/fixtures/hyphen-rules.py [out dir]   (default tests/word/fixtures/hyphen-rules)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/hyphen-rules \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/hyphen-rules --refresh-word
    python3 tools/word-parity/fixtures/long-word.py --lines tests/word/fixtures/hyphen-rules
    python3 tools/word-parity/fixtures/hyphen-rules.py --table tests/word/fixtures/hyphen-rules

--table prints, per probe and document, what Word did: "-<n>" broken after n letters of the probed word
plus a hyphen, "=" moved whole to the next line. For the flow documents (flow-*) it compares Word's lines
with the lines a greedy breaker makes from Word's own points (tests/word/fixtures/hyphen-words), under
each rule for the zone and the limit.

Courier New 10pt (6pt a character), lines exactly 20pt, a Letter page with 1in margins and a right
indent leaving 363pt: a line holds 60 characters, not 61. A probe is a paragraph: a token of n
characters ("<label>aaaa..."), a space and the probed word (WORD, usually "accessibility", which Word
breaks ac-ces-si-bil-ity: after 2, 5, 7, 10 letters), then " end". Its room k = 59 - n: the characters
left for WORD on the first line. For accessibility every k from 3 to 12 has a point that fits.

The documents (settings) x probes:
  z<zone>-m<mode>-<jc>   zone: none (no w:hyphenationZone), 1, 180, 360, 720, 1440 twips; compatibility mode
                         14 or 15; left or both (justified). Probes k03 ... k12 (accessibility).
  base                   zone none, mode 15, left - the probes below:
    k<kk>                accessibility at room kk
    caps / capsx         ACCESSIBILITY / Accessibility, room 12 and 6
    dig                  accessibil1ty, access1bility, accessibility2 (room 12)
    punct                "accessibility," "(accessibility)" "accessibility's" "accessibility." (room 12)
    nop                  w:noProof on the word's run
    gb / de / fr / none  w:lang en-GB / de-DE / fr-FR on the run; "none": no w:lang anywhere (the
                         document's styles have none)
    sup                  w:suppressAutoHyphens on the paragraph
    supstyle             a paragraph style with w:suppressAutoHyphens
    br                   a line of "aaaa", w:br, then the probe (the word after a manual line break)
    short                five-letter words (ab-sence ... ) room 3, 4: "begin" "basic" "users" "rural" "laser"
  caps-dnh               base's caps probes with w:doNotHyphenateCaps
  limit<n>               flow paragraph (accessibility over and over, every line ending in it) with
                         w:consecutiveHyphenLimit n = 1, 2; limit0 none
  bottom-m<mode>[-allow][-last]  f01 ... f30 one-line paragraphs, then a probe on the 31st and 32nd lines; -last: f31 too, the
                         probe's first line the page's last; mode 14 or 15; -allow: compat allowHyphenationAtTrackBottom
  colbottom-m15          two columns, the probe on the first column's last line
  squeeze-m<mode>[-nohyph]  justified lines of 5 or 10 digit tokens (4, 9 spaces), 61 ... 64 characters, then
                         " end": does a justified line take more than its width by narrowing its spaces?
  flow-m<mode>-<jc>-z<zone>  paragraphs of the corpus's words (seeded), every word with Word's measured points
No dependencies: the package is written as XML.
"""
import os
import random
import re
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
IND = '<w:ind w:right="2100"/>'
COLS = 60
WORD = 'accessibility'
PTS = {}   # Word's points (hyphen-words), read on use


def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def r(text, rpr=''):
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def token(label, k):
    n = COLS - 1 - k
    return (label + 'a' * n)[:n] if len(label) < n else label


def probe(label, k, word=WORD, rpr='', ppr='', jc=''):
    jcx = f'<w:jc w:val="{jc}"/>' if jc else ''
    return f'<w:p><w:pPr>{ppr}{IND}{jcx}</w:pPr>{r(token(label, k) + " ")}{r(word, rpr)}{r(" end")}</w:p>'


def styles(lang='<w:lang w:val="en-US"/>'):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
            '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" '
            f'w:eastAsia="Courier New"/><w:sz w:val="20"/><w:szCs w:val="20"/>{lang}</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="400" '
            'w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="NoHyph"><w:name w:val="No Hyph"/><w:basedOn w:val="Normal"/>'
            '<w:pPr><w:suppressAutoHyphens/></w:pPr></w:style></w:styles>')


def settings(zone=None, mode=15, limit=None, caps=False, allow=False):
    z = f'<w:hyphenationZone w:val="{zone}"/>' if zone is not None else ''
    lim = f'<w:consecutiveHyphenLimit w:val="{limit}"/>' if limit is not None else ''
    dnh = '<w:doNotHyphenateCaps/>' if caps else ''
    al = '<w:allowHyphenationAtTrackBottom/>' if allow else ''
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:autoHyphenation/>'
            f'{lim}{z}{dnh}<w:compat>{al}'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


def sect(cols=1):
    c = f'<w:cols w:num="{cols}" w:space="360"/>' if cols > 1 else ''
    return (f'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
            f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{c}</w:sectPr>')


TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, st=None, se=None, sc=None):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sc or sect()}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', st or styles())
        z.writestr('word/settings.xml', se or settings())


def kprobes(jc=''):
    return ''.join(probe(f'k{k:02d}', k, jc=jc) for k in range(3, 13))


def base_body():
    b = kprobes()
    b += probe('caps12', 12, 'ACCESSIBILITY') + probe('caps06', 6, 'ACCESSIBILITY')
    b += probe('capsx12', 12, 'Accessibility') + probe('capsx06', 6, 'Accessibility')
    b += probe('dig1', 12, 'accessibil1ty') + probe('dig2', 12, 'access1bility') + probe('dig3', 12, 'accessibility2')
    b += probe('pun1', 12, 'accessibility,') + probe('pun2', 12, '(accessibility)')
    b += probe('pun3', 12, "accessibility's") + probe('pun4', 12, 'accessibility.')
    b += probe('nop', 12, rpr='<w:noProof/>')
    b += probe('gb', 12, rpr='<w:lang w:val="en-GB"/>') + probe('de', 12, rpr='<w:lang w:val="de-DE"/>')
    b += probe('fr', 12, rpr='<w:lang w:val="fr-FR"/>')
    b += probe('sup', 12, ppr='<w:suppressAutoHyphens/>')
    b += probe('supstyle', 12, ppr='<w:pStyle w:val="NoHyph"/>')
    b += (f'<w:p><w:pPr>{IND}</w:pPr>{r("aaaa")}<w:r><w:br/></w:r>{r(token("br", 12) + " ")}{r(WORD)}{r(" end")}</w:p>')
    for w in ['absence', 'begin', 'basic', 'users', 'rural', 'laser']:
        for k in (3, 4):
            b += probe(f's{w[:3]}{k}', k, w)
    return b


def flow_accessibility():
    """One paragraph, every line ending in accessibility at room 12 (when the line before was hyphenated
    at 10: 'ity' + ' ' + a token of 44 + ' ' + 'accessibil-')."""
    parts = ['L00' + 'a' * 45, WORD]
    for i in range(1, 20):
        parts.append(f'L{i:02d}' + 'a' * 41)
        parts.append(WORD)
    return ' '.join(parts)


def squeeze(spaces, extra):
    """A line of spaces + 1 tokens of digits (not hyphenated), 60 + extra characters, then ' end'."""
    total = COLS + extra - spaces
    n = spaces + 1
    sizes = [total // n + (1 if i < total % n else 0) for i in range(n)]
    return ' '.join(str(i % 10) * size for i, size in enumerate(sizes)) + ' end'


def flow_words(seed, n=220):
    words = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../../tests/word/fixtures/hyphen-words/words.txt')).read().split()
    rnd = random.Random(seed)
    return ' '.join(rnd.choice(words) for _ in range(n))


def generate(out):
    os.makedirs(out, exist_ok=True)
    for f in os.listdir(out):
        if f.endswith('.docx'):
            os.remove(os.path.join(out, f))
    for zone in [None, 1, 180, 360, 720, 1440]:
        for mode in (14, 15):
            for jc in ('left', 'both'):
                name = f'z{zone if zone is not None else "none"}-m{mode}-{jc}'
                write(os.path.join(out, name + '.docx'), kprobes('' if jc == 'left' else 'both'), se=settings(zone, mode))
    write(os.path.join(out, 'base.docx'), base_body())
    write(os.path.join(out, 'base-m14.docx'), base_body(), se=settings(None, 14))
    write(os.path.join(out, 'none.docx'), base_body(), st=styles(lang=''))
    caps = (probe('caps12', 12, 'ACCESSIBILITY') + probe('caps06', 6, 'ACCESSIBILITY')
            + probe('capsx12', 12, 'Accessibility') + probe('capsx06', 6, 'Accessibility') + probe('k12', 12))
    write(os.path.join(out, 'caps-dnh.docx'), caps, se=settings(caps=True))
    for limit in (None, 0, 1, 2):
        name = 'limit' + ('none' if limit is None else str(limit))
        write(os.path.join(out, name + '.docx'), f'<w:p><w:pPr>{IND}</w:pPr>{r(flow_accessibility())}</w:p>',
              se=settings(limit=limit))
    fill = ''.join(f'<w:p><w:pPr>{IND}</w:pPr>{r(f"f{i:02d}")}</w:p>' for i in range(1, 31))
    for mode in (14, 15):
        for allow in (False, True):
            name = f'bottom-m{mode}' + ('-allow' if allow else '')
            # f01-f30, a probe (lines 31-32: its first line the 31st); in -last f01-f31 and the probe on the
            # 32nd line, the page's last (its second line on page 2).
            body = fill + probe('p31', 12)
            write(os.path.join(out, name + '-last.docx'), fill + f'<w:p><w:pPr>{IND}</w:pPr>{r("f31")}</w:p>' + probe('p32', 12),
                  se=settings(mode=mode, allow=allow))
            write(os.path.join(out, name + '.docx'), body, se=settings(mode=mode, allow=allow))
    # Two columns of 180pt... 60 characters need 363pt: the column is the line (no right indent), 4in each
    # needs a wider page: 11in wide, 0.5in margins, 2 columns of 363pt (7260 twips), 360 apart.
    col = (f'<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="720" '
           'w:bottom="1440" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
           '<w:cols w:num="2" w:space="360" w:equalWidth="0"><w:col w:w="7290" w:space="360"/><w:col w:w="7290"/></w:cols></w:sectPr>')
    colfill = ''.join(f'<w:p>{r(f"f{i:02d}")}</w:p>' for i in range(1, 23))

    def cprobe(label, k):
        return f'<w:p>{r(token(label, k) + " ")}{r(WORD)}{r(" end")}</w:p>'
    # (23 lines a column: f01-f22, the probe on the first column's last line, its second line atop the next.)
    write(os.path.join(out, 'colbottom-m15.docx'), colfill + cprobe('p23', 12) + cprobe('p25', 12), sc=col)
    for mode in (14, 15):
        for hyph in (True, False):
            body = ''.join(f'<w:p><w:pPr>{IND}<w:jc w:val="both"/><w:spacing w:after="400"/></w:pPr>{r(squeeze(sp, e))}</w:p>'
                           for sp in (4, 9) for e in (1, 2, 3, 4))
            se = settings(mode=mode)
            if not hyph:
                se = se.replace('<w:autoHyphenation/>', '')
            write(os.path.join(out, f'squeeze-m{mode}' + ('' if hyph else '-nohyph') + '.docx'), body, se=se)
    for mode in (14, 15):
        for jc in ('left', 'both'):
            for zone in (None, 1):
                name = f'flow-m{mode}-{jc}-z{"none" if zone is None else zone}'
                jcx = '<w:jc w:val="both"/>' if jc == 'both' else ''
                body = ''.join(f'<w:p><w:pPr>{IND}{jcx}<w:spacing w:after="400"/></w:pPr>{r(flow_words(s))}</w:p>' for s in (1, 2, 3))
                write(os.path.join(out, name + '.docx'), body, se=settings(zone, mode))
    print(len([f for f in os.listdir(out) if f.endswith('.docx')]), 'documents')


def load_points():
    if not PTS:
        f = os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../../tests/word/fixtures/hyphen-words/word/points.tsv')
        for line in open(f):
            c = line.rstrip('\n').split('\t')
            if c[0] == 'word':
                PTS[c[1]] = [int(x) for x in c[2].split(',') if x]
    return PTS


def greedy(text, width=COLS, limit=0, zone_chars=None):
    """Lines of text broken greedily, words hyphenated at Word's last point that fits (and the hyphen).
    zone_chars: a word is hyphenated only when the room it leaves without hyphenating exceeds it."""
    pts = load_points()
    lines, cur, run = [], '', 0
    for w in text.split(' '):
        while True:
            add = (' ' if cur else '') + w
            if len(cur) + len(add) <= width:
                cur += add
                break
            room = width - len(cur) - (1 if cur else 0)
            core = re.sub(r'[^a-z]+$', '', w)
            best = 0
            if (limit == 0 or run < limit) and (zone_chars is None or width - len(cur) > zone_chars):
                best = max([p for p in pts.get(core, []) if p + 1 <= room], default=0)
            if best:
                lines.append(cur + (' ' if cur else '') + w[:best] + '-')
                run += 1
                w = w[best:]
                cur = ''
            else:
                lines.append(cur)
                run = 0
                cur = ''
    lines.append(cur)
    return lines


def table(folder):
    rows = [l.split('\t') for l in open(os.path.join(folder, 'word', 'lines.tsv'), encoding='utf-8').read().splitlines()]
    by = {}
    for name, page, top, text in rows:
        by.setdefault(name, []).append(text)
    for name in sorted(by):
        ls = by[name]
        if name.startswith('flow') or name.startswith('limit'):
            word_lines = ls
            print(name, len(word_lines), 'lines')
            continue
        res = []
        for i, t in enumerate(ls):
            m = re.match(r'^([a-z]+\d+)a* (\S+)?$', t)
            if m and not t.endswith(' end'):
                word = m.group(2) or ''
                res.append(f'{m.group(1)}:' + (f'-{len(word) - 1}' if word.endswith('-') else ('=' if not word else '?' + word)))
            elif re.match(r'^[a-z]+\d*a* \S+ end$', t):
                res.append(t.split(' ')[0][:8] + ':whole')
        print(f'{name:24s} ' + ' '.join(res))


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--table':
        table(sys.argv[2])
        sys.exit(0)
    generate(sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/hyphen-rules'))
