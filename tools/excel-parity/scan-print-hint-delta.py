#!/usr/bin/env python3
"""Scans fonts for ppem where FreeType-v35 hinting differs from design rounding.

This is the selection step for the second challenge: only faces/ppem with a
nonzero hinted-vs-baseline delta can discriminate.  It uses the same ctypes
FreeType bridge and the same characters as the print oracle.

    python3 tools/excel-parity/scan-print-hint-delta.py OUT.tsv
"""
import importlib.util
import math
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('lockhint', os.path.join(_here, 'lock-print-hint-oracle.py'))
lockhint = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(lockhint)

FACES = {
    'Arial': 'artifacts/fonts/windows/arial.ttf',
    'Times New Roman': 'artifacts/fonts/windows/times.ttf',
    'Consolas': 'artifacts/fonts/windows-extra/consola.ttf',
    'Lucida Console': 'artifacts/fonts/windows-extra/lucon.ttf',
    'Garamond': 'artifacts/fonts/windows/GARA.TTF',
    'Tahoma': 'artifacts/fonts/windows-extra/tahoma.ttf',
}
CHARS = '0123456789.,%$()-#'
PPEMS = range(40, 221)


def main():
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    out = ['face\tppem\tchar\tdesign_dots\thinted_dots\tdelta']
    for name, rel in FACES.items():
        path = os.path.join(lockhint.gen.ROOT, rel)
        ems = lockhint.design_metrics(path)
        for ppem in PPEMS:
            for ch in CHARS:
                if ch not in ems:
                    continue
                design = math.floor(ems[ch] * ppem + 0.5)
                hinted = lockhint.hinted_px(path, ppem, ch)
                if hinted != design:
                    out.append(f'{name}\t{ppem}\t{ch}\t{design}\t{hinted:g}\t{hinted - design:+g}')
    with open(sys.argv[1], 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} delta rows to {sys.argv[1]}')


if __name__ == '__main__':
    main()
