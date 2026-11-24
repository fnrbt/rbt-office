#!/usr/bin/env python3
"""Key-set verifier for the border/fill discriminating fixture (multipage-safe).

    python3 tools/excel-parity/verify-print-border-fill-keys.py CELLS.tsv MEASURED.tsv [MEASURED.tsv ...]

Expected keys are every (sheet, row, width) the manifest declares. The 100 % border workbooks
paginate their final width column horizontally, so a verifier that silently accepts the first page
would hide 144 expected keys per workbook (12 sheets x 12 rows). Uses the same mandatory
check_key_sets gate as measure-print-border-fill.py. Exits 1 on any missing, extra or duplicate key.
"""
import collections
import importlib.util
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location(
    'borderfill', os.path.join(HERE, 'measure-print-border-fill.py'))
bf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bf)


def load(path):
    lines = [l.rstrip('\n') for l in open(path, encoding='utf-8') if l.strip()]
    head = lines[0].split('\t')
    return [dict(zip(head, l.split('\t'))) for l in lines[1:]]


def main():
    if len(sys.argv) < 3:
        raise SystemExit('usage: verify-print-border-fill-keys.py CELLS.tsv MEASURED.tsv [...]')
    expected = bf.expected_keys(sys.argv[1])
    bad = False
    for path in sys.argv[2:]:
        actual = collections.Counter(
            (r['sheet'], r['row'], r['width']) for r in load(path))
        bad = not bf.check_key_sets(expected, actual, path) or bad
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
