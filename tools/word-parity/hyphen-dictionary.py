#!/usr/bin/env python3
"""Writes src/Word/HyphenWordsEnUs.fs: the Moby hyphenation list's points for common words where they
differ from the patterns (hyph-en-us, src/Word/HyphenPatternsEnUs.fs). docs/layout-parity-flow.md, ledger.

    python3 tools/word-parity/hyphen-dictionary.py <mhyph.txt> <count_1w.txt> [top]   (top: 50000)

mhyph.txt: Grady Ward's Moby Hyphenator II (public domain; Project Gutenberg etext 3204,
https://gutenberg.pglaf.org/3/2/0/3204/files/mhyph.txt), its points marked with 0xA5.
count_1w.txt: Peter Norvig's word counts of Google's web corpus (https://norvig.com/ngrams/count_1w.txt),
only to choose the words: the `top` most frequent ones. Nothing of it is written out.

Word's points are a dictionary's: on 2000 words checked in Word that the measured sample did not have
(tests/word/fixtures/hyphen-words-new), Moby gives Word's points for 98.6% of those it lists, the patterns
for 87.7%. An entry is a lower-case word of 5 letters or more, its points at least 2 letters in and 2 from
its end (the layout's own minimums apply on top: 3 from the end in Word 2013's layout).
"""
import os
import re
import sys

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '../..'))


def fsharp_strings(text):
    return ' '.join(re.findall(r'"([^"]*)"', text))


def load_patterns():
    t = open(os.path.join(REPO, 'src/Word/HyphenPatternsEnUs.fs'), encoding='utf-8').read()
    pi, ei = t.index('let patterns'), t.index('let exceptions')
    values, longest = {}, 0
    for p in fsharp_strings(t[pi:ei]).split():
        letters, digits = '', [0]
        for ch in p:
            if ch.isdigit():
                digits[-1] = int(ch)
            else:
                letters += ch
                digits.append(0)
        values[letters] = digits
        longest = max(longest, len(letters))
    exceptions = {}
    for e in fsharp_strings(t[ei:]).split():
        exceptions[e.replace('-', '')] = breaks(e)
    return values, longest, exceptions


def breaks(entry):
    out, n = [], 0
    for ch in entry:
        if ch == '-':
            out.append(n)
        else:
            n += 1
    return out


def pattern_points(values, longest, exceptions, word):
    """As Hyphenation.pointsOf with 2 letters in and 2 from the end."""
    n = len(word)
    if word in exceptions:
        return [p for p in exceptions[word] if p >= 2 and n - p >= 2]
    s = '.' + word + '.'
    v = [0] * (len(s) + 1)
    for i in range(len(s)):
        for j in range(i + 1, min(len(s), i + longest) + 1):
            d = values.get(s[i:j])
            if d:
                for k, x in enumerate(d):
                    if x > v[i + k]:
                        v[i + k] = x
    return [p for p in range(2, n - 1) if v[p + 1] % 2 == 1]


def word_points():
    """Word's points of the words of the measured samples (mode 15 and mode 14 alike)."""
    out = {}
    base = os.path.join(REPO, 'tests/word/fixtures')
    for d in sorted(os.listdir(base)):
        f = os.path.join(base, d, 'word', 'points.tsv')
        if not d.startswith('hyphen-words') or not os.path.exists(f):
            continue
        for line in open(f, encoding='utf-8'):
            c = line.rstrip('\n').split('\t')
            if c[0] == 'word' and len(c) >= 3 and re.fullmatch(r'[a-z]+', c[1]):
                out.setdefault(c[1], set()).update(int(x) for x in c[2].split(',') if x)
    return out


def main():
    moby_file, count_file = sys.argv[1], sys.argv[2]
    top = int(sys.argv[3]) if len(sys.argv) > 3 else 50000
    common = set()
    for line in open(count_file, encoding='utf-8'):
        if len(common) >= top:
            break
        common.add(line.split('\t')[0])
    values, longest, exceptions = load_patterns()
    entries = {}
    raw = open(moby_file, 'rb').read().decode('latin-1').replace('\xa5', '-')
    for line in raw.splitlines():
        e = line.strip()
        if not re.fullmatch(r'[a-z]+(-[a-z]+)*', e):
            continue
        w = e.replace('-', '')
        if len(w) < 5 or w not in common or w in entries:
            continue
        pts = [p for p in breaks(e) if p >= 2 and len(w) - p >= 2]
        if pts != pattern_points(values, longest, exceptions, w):
            entries[w] = pts
    # Word's own points, where Word was asked (the measured samples: tests/word/fixtures/hyphen-words*,
    # word/points.tsv): over Moby's, which differ from Word's for 1.4% of the words (serv-ice, Word ser-vice).
    # (A point 2 letters from the end, which Word 2013's layout never shows, as Moby or the patterns have it.)
    moby = {}
    for line in raw.splitlines():
        e = line.strip()
        if re.fullmatch(r'[a-z]+(-[a-z]+)*', e):
            moby.setdefault(e.replace('-', ''), [p for p in breaks(e) if p >= 2])
    for w, seen in word_points().items():
        base = moby.get(w) or pattern_points(values, longest, exceptions, w)
        pts = sorted(set(seen) | {p for p in base if len(w) - p == 2})
        pts = [p for p in pts if p >= 2 and len(w) - p >= 2]
        if pts != pattern_points(values, longest, exceptions, w):
            entries[w] = pts
        else:
            entries.pop(w, None)

    def show(w, pts):
        return ''.join(('-' if i in pts else '') + ch for i, ch in enumerate(w))
    items = [show(w, entries[w]) for w in sorted(entries)]
    lines, cur = [], ''
    for it in items:
        if cur and len(cur) + 1 + len(it) > 110:
            lines.append(cur)
            cur = it
        else:
            cur = (cur + ' ' + it) if cur else it
    if cur:
        lines.append(cur)
    out = os.path.join(REPO, 'src/Word/HyphenWordsEnUs.fs')
    with open(out, 'w', encoding='utf-8') as f:
        f.write(f'''/// The points of the Moby hyphenation list (Grady Ward's Moby Hyphenator II, public domain) for the
/// {top} most frequent words of the web (Norvig's count_1w) where they differ from the patterns
/// (HyphenPatternsEnUs.fs): Word's points are a dictionary's, and Moby's are Word's far more often than the
/// patterns' (docs/layout-parity-flow.md, ledger). Generated by tools/word-parity/hyphen-dictionary.py;
/// installed after the patterns' exceptions (Hyphenation.fs), loaded on use with them.
module Rbt.Word.HyphenWordsEnUs

/// {len(items)} words, their points marked "-" (at least 2 letters in and 2 from the end; none: not broken).
let words =
    String.concat " " [
''')
        for l in lines:
            f.write(f'        "{l}"\n')
        f.write('    ]\n')
    print(len(items), 'words written to', out)


if __name__ == '__main__':
    main()
