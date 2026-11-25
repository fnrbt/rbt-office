#!/usr/bin/env python3
"""Cases for the regular expression engine (src/Spreadsheet/Functions/Pcre.fs), with perl as the oracle.

    python3 tools/excel-parity/make-regex-cases.py > tools/excel-parity/regex-cases.jsonl
    perl tools/excel-parity/regex-perl-oracle.pl < tools/excel-parity/regex-cases.jsonl > tools/excel-parity/regex-perl.jsonl
    dotnet fsi tools/excel-parity/regex-compare.fsx

One JSON object per line: pattern, subject, i (caseless). Hand-written cases of every construct Excel
users meet (classes, quantifiers greedy / lazy / possessive, groups, names, backreferences, lookaround,
anchors, \\b, inline options, Unicode, conditionals, branch reset), then random patterns of a small
grammar over a small alphabet (fixed seed) to shake out the backtracking. Perl is not PCRE2: the cases
keep to what the two agree on (\\w and the POSIX classes ASCII: the oracle runs with /a; no \\X, no
recursion that backtracks into a call, no (*UCP)).
"""
import json
import random

cases = []


def case(pattern, *subjects, i=False):
    for s in subjects:
        cases.append({'pattern': pattern, 'subject': s, 'i': i})


# literals, classes, escapes
case('a', 'alfalfa', 'xyz', '')
case('[a-z]', 'alfalfa', 'ABC')
case('[A-Z]', 'alfalfa', 'aBc')
case('[aeiou]', 'alfalfa', 'rhythm')
case('[0-9]', 'alfalfa', 'a1')
case('^\\([0-9]{3}\\) [0-9]{3}-[0-9]{4}$', '(378) 555-4195', '+1(878) 555-8622')
case('[A-Z][a-z]+', 'DylanWilliams', 'SoniaBrown')
case('[0-9()]+ [0-9-]+', 'Sonia Rees (378) 555-4195\nAngel Brown (878) 555-8622\nBlake Martin (437) 555-8987')
case('[0-9]+-', 'Sonia Rees(378) 555-4195\nAngel Brown(878) 555-8622')
case('([A-Z][a-z]+)([A-Z][a-z]+)', 'SoniaBrown')
case('[^0-9]', 'a1b2', '123')
case('[]a]', ']', 'a', 'b')
case('[^]a]', ']', 'b')
case('[a-]', '-', 'b')
case('[-a]', '-', 'b')
case('[\\d-]+', '12-34x')
case('[\\w.]+@[\\w.]+', 'mail me at a.b@c.org now')
case('\\d+', 'abc 123 def 45')
case('\\D+', 'abc 123')
case('\\w+', 'héllo wörld_1')
case('\\W+', 'héllo wörld')
case('\\s+', 'a \t\nb')
case('\\S+', ' ab  cd ')
case('\\h+', 'a \t b')
case('\\v', 'a\nb', 'ab')
case('\\R', 'a\r\nb', 'a\rb', 'ab')
case('\\N+', 'ab\ncd')
case('[[:alpha:]]+', 'abc123')
case('[[:digit:][:space:]]+', 'x1 2y')
case('[[:^alpha:]]+', 'ab12cd')
case('[[:punct:]]+', 'a,.;b')
case('[[:upper:]][[:lower:]]+', 'heLLo World')
case('[[:xdigit:]]+', 'zz0fA9g')
case('\\x41\\x{42}\\103', 'ABC')
case('\\t\\n', 'a\t\nb')
case('\\cA', '\x01')
case('\\e', '\x1b')
case('\\o{101}', 'A')
case('\\N{U+00E9}', 'café')
case('\\Q.*\\E+', '.*.*', 'abc')
case('a\\.b', 'a.b', 'axb')
case('\\$\\d', 'cost $5')
case('[\\x{100}-\\x{17F}]+', 'žluťoučký')
# quantifiers
case('a*', 'aaa', 'baaa', '')
case('a+', 'baaa', 'b')
case('a?', 'ab', 'b')
case('a{2}', 'aaaaa')
case('a{2,}', 'aaaaa', 'a')
case('a{2,3}', 'aaaaaaa')
case('a{,2}', 'aaa')
case('a{,}', 'a{,}')
case('x{', 'x{')
case('x{a}', 'x{a}')
case('a*?', 'aaa')
case('a+?', 'aaa')
case('a??b', 'ab', 'b')
case('a{2,3}?', 'aaaa')
case('a*+a', 'aaaa')
case('a++b', 'aaab')
case('a?+a', 'a', 'aa')
case('(a|ab)(c|bcd)(d*)', 'abcd')
case('<.+>', '<a><b>')
case('<.+?>', '<a><b>')
case('(ab)+', 'ababab')
case('(a*)*', 'aaa', 'b')
case('(a*)+', 'aaa', 'b')
case('(a|b)*c', 'ababc')
case('(?:a|)*b', 'aab')
case('(a?)*?b', 'aab')
case('(x+x+)+y', 'xxxxxxxxxxy', 'xxxxxy')
# groups, names, backreferences
case('(\\d+)-(\\d+)', 'tel 555-1234')
case('(?<year>\\d{4})-(?<month>\\d\\d)', 'on 2024-05-17')
case("(?'w'\\w+) (?P<x>\\w+)", 'hello world')
case('(\\w)\\1', 'hello', 'abc')
case('(\\w)(\\w)\\2\\1', 'abba xyyx')
case('(?<c>.)\\k<c>', 'aabb')
case('(?<c>.)\\k{c}', 'xyy')
case("(?<c>.)\\k'c'", 'xyy')
case('(?P<c>.)(?P=c)', 'xyy')
case('(a)\\g1', 'aa')
case('(a)\\g{1}', 'aa')
case('(a)(b)\\g{-1}', 'abb')
case('(a)|b', 'b', 'a')
case('(a)?b', 'b', 'ab')
case('(?:(a)|(b))+', 'ab', 'ba')
case('(a)|(b)|(c)', 'c')
case('(?|(a)|(b))', 'b')
case('(?|(a)(x)|(b))c', 'bc', 'axc')
case('((a)(b))', 'ab')
case('(a(b(c)))', 'abc')
case('(?:ab)+', 'abab')
case('(?>a+)b', 'aaab')
case('(?>a+)a', 'aaa')
case('(?>ab|a)b', 'ab', 'abb')
# lookaround
case('foo(?=bar)', 'foobar', 'foobaz')
case('foo(?!bar)', 'foobar', 'foobaz')
case('(?<=\\$)\\d+', 'cost $42 or 7')
case('(?<!\\$)\\b\\d+', 'cost $42 or 7')
case('(?<=a|bc)x', 'ax', 'bcx', 'cx')
case('(?<=ab|c)d', 'abd', 'cd', 'bd')
case('\\b\\w+(?=ing\\b)', 'singing and dancing')
case('(?=(\\d{3}))', '12345')
case('(?=.*\\d)(?=.*[a-z]).{6,}', 'abc123', 'abcdef', '12345678a')
case('(?<=(a))b', 'ab')
case('(?!a)\\w', 'ab')
case('q(?=u)', 'quit', 'qatar')
case('(?<![a-z])\\d', 'a1 2')
# anchors, \b, multiline
case('^abc', 'abc', 'xabc')
case('abc$', 'abc', 'abc\n', 'abcx')
case('^$', '', '\n', 'a')
case('$', 'ab\n', 'ab')
case('^', 'a\nb\n')
case('(?m)^\\w', 'ab\ncd\n')
case('(?m)\\w$', 'ab\ncd\n')
case('(?m)^$', 'a\n\nb\n')
case('\\Aab', 'ab', 'cab')
case('ab\\z', 'ab', 'ab\n')
case('ab\\Z', 'ab', 'ab\n', 'ab\n\n')
case('\\bcat\\b', 'cat', 'concat', 'the cat sat')
case('\\Bcat\\B', 'concatenate', 'cat')
case('\\b', 'ab cd')
case('\\B', 'ab cd')
case('\\Gab', 'ababxab')
case('a\\Kb', 'ab')
case('foo\\Kbar', 'foobar')
case('(?s).+', 'a\nb')
case('.+', 'a\nb')
case('(?s:.)\\.', '\n.')
# inline options and case
case('(?i)abc', 'ABC', 'aBc')
case('abc', 'ABC', 'abc', i=True)
case('a(?i)b', 'aB', 'AB')
case('a(?i:b)c', 'aBc', 'aBC')
case('(?i)a(?-i)b', 'Ab', 'AB')
case('(a(?i)b|c)', 'aB', 'C')
case('[a-c]+', 'ABCabc', i=True)
case('[^a]', 'A', i=True)
case('(\\w)\\1', 'aA', i=True)
case('(?x) a b # comment\n c', 'abc')
case('(?x)[ ]a', ' a')
case('(?xx)[a b]+', 'ab b')
case('(?n)(a)(?<x>b)', 'ab')
case('(?U)a+', 'aaa')
case('(?U)a+?', 'aaa')
case('(?^i:a)A', 'aA')
case('straße', 'STRASSE', 'STRAßE', i=True)
case('k', 'K', 'K', i=True)
case('σ', 'Σ', 'ς', i=True)
case('é', 'É', i=True)
case('[à-ÿ]+', 'ÀÉÎ', i=True)
case('ǅ', 'Ǆ', 'ǆ', i=True)
# Unicode
case('\\p{L}+', 'héllo wörld 123')
case('\\p{Lu}', 'abcDef')
case('\\p{Ll}+', 'ABcdEF')
case('\\pL+', 'αβγ abc')
case('\\P{L}+', 'ab12cd')
case('\\p{^L}+', 'ab12cd')
case('\\p{N}+', 'x٣٤5y')
case('\\p{Nd}', '٣')
case('\\p{Greek}+', 'abc αβγ')
case('\\p{Han}+', '漢字かな')
case('\\p{Hiragana}+', '漢字かな')
case('\\p{Katakana}+', 'カタカナー')
case('\\p{Cyrillic}+', 'привет мир')
case('\\p{Latin}+', 'Ünïcödé 字')
case('\\p{Arabic}+', 'مرحبا')
case('\\p{Sc}', 'cost €5 or $3')
case('\\p{P}+', 'a,b!?')
case('\\p{Zs}', 'a b')
case('\\p{L&}+', 'abcǅ')
case('\\p{Lu}\\p{Ll}*', 'ÉCOLE École')
case('.', '😀', '')
case('..', '😀a')
case('[😀-😂]', '😁')
case('\\x{1F600}', '😀')
case('😀+', '😀😀x')
case('[^a]', '😀')
case('\\w+', 'naïve café')
case('[[:alpha:]]+', 'naïve')
case('\\b\\w+\\b', 'naïve')
case('ü', 'Ü', 'ü', i=True)
# conditionals, recursion-free calls
case('(a)?(?(1)b|c)', 'ab', 'c', 'b')
case('(?<q>")?\\w+(?(q)")', '"abc"', 'abc', '"abc')
case('(?(?=\\d)\\d{3}|[a-z]+)', '123', 'abc')
case('(?(?!\\d)[a-z]+|\\d)', '5', 'xy')
case('(?(DEFINE)(?<d>\\d))(?&d)+', 'ab12')
case('(\\d)(?1)', '12')
case('(?<n>[a-z])(?&n)', 'ab')
case('(a|b)c\\g<1>', 'acb')
case('\\((?:[^()]|(?R))*\\)', 'x(a(b)c)y')
case('a(*F)|b', 'ab')
case('(*pla:a)\\w', 'ab')
case('(*nlb:a)b', 'ab', 'cb')
case('(*atomic:a+)b', 'aab')
# empty matches in global iteration
case('', 'abc')
case('x*', 'axb')
case('\\b', 'a b')
case('(?=a)', 'aaa')
case('a|', 'bab')
case('|a', 'bab')
case('a?', 'baab')
case('(?=b)|b', 'abc')
# more realistic
case('[\\w.+-]+@[\\w-]+\\.[\\w.]+', 'contact: john.doe+x@mail.example.com, jane@ex.org')
case('https?://[^\\s]+', 'see https://example.com/a?b=c and http://x.org.')
case('\\d{1,3}(,\\d{3})*(\\.\\d+)?', 'total 1,234,567.89 units')
case('(\\d{4})-(\\d{2})-(\\d{2})', '2024-01-15 and 2023-12-31')
case('^\\s+|\\s+$', '  padded  ')
case('\\s{2,}', 'a  b     c')
case('(\\w+)\\s\\1', 'the the cat')
case('[A-Z]{2,}', 'NASA and ESA')
case('\\b[A-Z]', 'hello World Wide web')
case('(?i)\\bthe\\b', 'The other the THE')
case('(?<=\\s|^)#\\w+', '#start middle #tag end#no')
case('^(?:(?:25[0-5]|2[0-4]\\d|1?\\d?\\d)\\.){3}(?:25[0-5]|2[0-4]\\d|1?\\d?\\d)$', '192.168.1.1', '256.1.1.1')
case('([^,]+),?', 'a,b,,c')
case('"([^"]*)"', 'say "hi" and "bye"')
case('(\\w+)@(\\w+)\\.com', 'x@y.com z@w.com')
case('[^\\x00-\\x7F]+', 'abc漢字def')
case('\\$(\\d+)\\.(\\d{2})', 'price: $19.99')

# random patterns of a small grammar
R = random.Random(20261125)
ATOMS = ['a', 'b', 'c', '.', '[ab]', '[^a]', '\\w', '\\d', '\\s', '\\b', '^', '$', '(?=a)', '(?!b)', '(?<=a)', '(?<!b)',
         '\\1', '(?:a|b)', 'x']


def rand_pattern(depth=0):
    parts = []
    for _ in range(R.randint(1, 4)):
        r = R.random()
        if depth < 2 and r < 0.25:
            inner = rand_pattern(depth + 1)
            kind = R.choice(['(%s)', '(?:%s)', '(?>%s)', '(%s|%s)', '(?=%s)', '(?!%s)'])
            atom = kind % ((inner,) if kind.count('%s') == 1 else (inner, rand_pattern(depth + 1)))
        else:
            atom = R.choice(ATOMS)
        if atom not in ('\\b', '\\B', '^', '$') and not atom.startswith('(?=') and not atom.startswith('(?!') \
                and not atom.startswith('(?<') and R.random() < 0.45:
            atom += R.choice(['*', '+', '?', '{2}', '{1,3}', '*?', '+?', '??', '*+', '++', '{0,2}?'])
        parts.append(atom)
    return ''.join(parts)


subjects = ['', 'a', 'ab', 'aab', 'abc', 'abab', 'ba', 'a b', 'aa bb', 'xa1', 'abca', 'cab ab', 'a\nb', 'aaaa']
n = 0
while n < 1500:
    p = rand_pattern()
    if '\\1' in p and '(' not in p.split('\\1')[0].replace('(?', ''):
        continue
    for s in R.sample(subjects, 3):
        cases.append({'pattern': p, 'subject': s, 'i': R.random() < 0.15})
    n += 1

# a second grammar: case, word boundaries, lookbehind of several lengths, two groups referred back to
ATOMS = ['a', 'A', 'b', '[a-c]', '[^ab]', '\\W', '\\B', '\\b', '(?<=ab|c)', '(?<!a|bc)', '(?i:a)', '\\1', '\\2',
         '(a|ab)', '(b*)', '(?:|a)', 'c', ' ', '$', '(?m:^)']
subjects = ['', 'Aa', 'ab Ab', 'abcab', 'bcbc a', 'aAbB', 'cab\nab', 'ab ab ab', 'cc  aa']
n = 0
while n < 800:
    p = rand_pattern()
    for s in R.sample(subjects, 3):
        cases.append({'pattern': p, 'subject': s, 'i': R.random() < 0.3})
    n += 1

for c in cases:
    print(json.dumps(c, ensure_ascii=False))
