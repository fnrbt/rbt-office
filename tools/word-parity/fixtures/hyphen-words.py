#!/usr/bin/env python3
"""Fixtures for where Word hyphenates English words (w:autoHyphenation; docs/layout-parity-flow.md).

    python3 tools/word-parity/fixtures/hyphen-words.py [out dir]   (default tests/word/fixtures/hyphen-words)

then, to ask Word (winvm):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/hyphen-words \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/hyphen-words --refresh-word
    python3 tools/word-parity/fixtures/hyphen-words.py --points tests/word/fixtures/hyphen-words

--points reads Word's PDFs (<dir>/word/<name>.pdf, not committed) and writes <dir>/word/points.tsv: for
every word, the points Word broke it at (letters before each hyphen, comma-separated) and, per probe,
what it did. The words: <dir>/words.txt (8092 lower-case words of 5-20 letters: every such word of the
corpus's documents and the common English words of 5 letters and more of a frequency list that are in
dwyl/english-words (The Unlicense)).

Every probe forces Word to decide: Courier New 6pt (3.6pt a character, no kerning), lines exactly 7.2pt,
four columns of 33 characters (not 34), left aligned, no widow control, en-US, w:autoHyphenation with a
1pt hyphenation zone and no limit on consecutive hyphens. A probe is a paragraph: a label of digits
(the probe's number, zero-padded) filling the column but for a space and k characters, then the word
(n letters, k = 3 ... n-1). The word does not fit whole; Word either breaks it at a point p with
p letters and the hyphen in the k characters, or moves it whole to the next line. If Word breaks at the
last point that fits, the probes of a word give all its points from 2 letters in to 1 letter from its end.
--c14 <dir> writes the same probes in compatibility mode 14 (Word 2010's layout) for <dir>/words.txt:
tests/word/fixtures/hyphen-words-c14 has the first 1000 words (the hyphenation of Word 2010's layout differs).
No dependencies: the package is written as XML.
"""
import os
import re
import subprocess
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
COLS = 33          # characters a column holds
PER_DOC = 4800     # probes a document


def styles(lang='en-US'):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
            '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" '
            f'w:eastAsia="Courier New"/><w:sz w:val="12"/><w:szCs w:val="12"/><w:lang w:val="{lang}"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="144" '
            'w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')


def settings(zone=20, limit=0, mode=15):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:autoHyphenation/>'
            f'<w:consecutiveHyphenLimit w:val="{limit}"/><w:hyphenationZone w:val="{zone}"/><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/>'
            '</w:compat></w:settings>')


# A column 33.5 characters wide (2412 twips), four of them 360 apart, 0.5in margins.
SECT = ('<w:sectPr><w:pgSz w:w="12168" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
        'w:header="360" w:footer="360" w:gutter="0"/><w:cols w:num="4" w:space="360"/></w:sectPr>')

TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''


def write(path, body, sect=SECT, st=None, se=None):
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{sect}</w:body></w:document>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', st or styles())
        z.writestr('word/settings.xml', se or settings())


def probes(words):
    """(number, word, k) for every word and room k."""
    out = []
    for w in words:
        for k in range(3, len(w)):
            out.append((len(out), w, k))
    return out


def label(i, k):
    return str(i).zfill(COLS - 1 - k)


def para(text):
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def generate(out_dir, mode=15):
    words = open(os.path.join(out_dir, 'words.txt')).read().split()
    ps = probes(words)
    for f in os.listdir(out_dir):
        if f.endswith('.docx'):
            os.remove(os.path.join(out_dir, f))
    for d in range(0, len(ps), PER_DOC):
        chunk = ps[d:d + PER_DOC]
        write(os.path.join(out_dir, f'w{d // PER_DOC + 1:02d}.docx'),
              ''.join(para(label(i, k) + ' ' + w) for i, w, k in chunk), se=settings(mode=mode))
    print(len(words), 'words,', len(ps), 'probes,', (len(ps) + PER_DOC - 1) // PER_DOC, 'documents')


def pdf_lines(pdf):
    """The lines of a PDF of four columns, in reading order: (page, column, top, text)."""
    out = subprocess.run(['pdftotext', '-bbox', pdf, '-'], capture_output=True, text=True, check=True).stdout
    rows = []
    for page, chunk in enumerate(out.split('<page ')[1:], 1):
        words = [(float(m.group(1)), float(m.group(2)), m.group(3)) for m in
                 re.finditer(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">(.*?)</word>', chunk)]
        # Columns start at 36pt and every 138.6pt (2412 + 360 twips) after.
        groups = {}
        for x, top, t in words:
            col = int((x - 36 + 10) // 138.6)
            key = (col, round(top / 3.6))
            groups.setdefault(key, []).append((x, t))
        for (col, topk), ws in sorted(groups.items()):
            rows.append((page, col, topk, ' '.join(t for _, t in sorted(ws))))
    return rows


def points(folder):
    words = open(os.path.join(folder, 'words.txt')).read().split()
    ps = probes(words)
    seen = {}          # probe number -> (break letters or 0, first-line's word part, rest)
    lines = []
    for name in sorted(f[:-5] for f in os.listdir(folder) if f.endswith('.docx')):
        pdf = os.path.join(folder, 'word', name + '.pdf')
        if not os.path.exists(pdf):
            print('no PDF for', name, file=sys.stderr)
            continue
        for page, col, top, text in pdf_lines(pdf):
            lines.append((name, page, col, text))
    cur = None
    for name, page, col, text in lines:
        m = re.fullmatch(r'(\d{12,})(?: (\S*))?', text)
        if m :
            cur = int(m.group(1))
            part = m.group(2) or ''
            seen[cur] = [part, '', page, col]
        elif cur is not None:
            seen[cur][1] += text
    rows, bad = [], 0
    by_word = {}
    for i, w, k in ps:
        if i not in seen:
            continue
        part, rest, page, col = seen[i]
        if part == '':
            got = 0
            ok = rest == w
        elif part.endswith('-'):
            got = len(part) - 1
            ok = part[:-1] + rest == w
        else:
            got, ok = -1, False
        if not ok:
            bad += 1
        by_word.setdefault(w, []).append((k, got))
        rows.append(f'probe\t{i}\t{w}\t{k}\t{got}\t{part}\t{rest}\t{page}\t{col}')
    out = []
    for w in words:
        if w not in by_word:
            continue
        pts = sorted({g for _, g in by_word[w] if g > 0})
        out.append(f'word\t{w}\t{",".join(map(str, pts))}')
    with open(os.path.join(folder, 'word', 'points.tsv'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(out + rows) + '\n')
    print(len(by_word), 'words read,', len(rows), 'probes,', bad, 'unreadable')


if __name__ == '__main__':
    repo = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..'))
    if len(sys.argv) > 2 and sys.argv[1] == '--points':
        points(sys.argv[2])
        sys.exit(0)
    if len(sys.argv) > 2 and sys.argv[1] == '--c14':
        generate(sys.argv[2], 14)
        sys.exit(0)
    generate(sys.argv[1] if len(sys.argv) > 1 else os.path.join(repo, 'tests/word/fixtures/hyphen-words'))
