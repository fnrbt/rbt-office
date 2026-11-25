#!/usr/bin/env python3
"""The formulas of tests/formula/stage70-regex-trimrange-text.fsx that need no cells (the REGEX*, ASC
and DBCS ones written against `empty`), one JSON string per line, for comparing the browser build with
.NET (regex-js-check.mjs, regex-formulas-net.fsx).

    python3 tools/excel-parity/make-regex-formulas.py > tools/excel-parity/regex-formulas.jsonl
"""
import json
import os
import re

here = os.path.dirname(os.path.abspath(__file__))
source = open(os.path.join(here, '..', '..', 'tests', 'formula', 'stage70-regex-trimrange-text.fsx'), encoding='utf-8').read()
literal = r'"((?:[^"\\]|\\.)*)"'
seen = []
for m in re.finditer(r'\bempty ' + literal, source):
    text = m.group(1)
    formula = re.sub(r'\\(.)', lambda e: {'n': '\n', 't': '\t'}.get(e.group(1), e.group(1)), text)
    if formula.startswith('=') and formula not in seen:
        seen.append(formula)
# the patterns PCRE2 rejects, written by a loop there
for bad in ["(", ")", "[a", "a{2,1}", "*a", "+", "a**", "(?<n>a)(?<n>b)", "\\", "[z-a]", "(?<=a+)b", "\\k<nope>",
            "\\9", "(?P<1a>x)", "\\x{110000}", "[[:nope:]]", "\\p{Nope}", "(?z)", "(*SKIP)a", "a(?#x", "\\c", "(?(1)a|b|c)",
            "\\i", "[\\d-z]", "x{70000}", "\\g{0}", "(?<=\\d+)"]:
    seen.append('=REGEXTEST("a","%s")' % bad.replace('"', '""'))
for f in seen:
    print(json.dumps(f, ensure_ascii=False))
