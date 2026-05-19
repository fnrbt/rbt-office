#!/usr/bin/env python3
"""Diff two VBA-style layout TSVs (Excel COM vs our spreadsheet engine).

Run `tools/excel-parity/excel-layout-tsv.sh foo.xlsx excel.tsv` to get
Excel's TSV, then `tools/excel-parity/spreadsheet-vba-layout-tsv.sh
foo.xlsx ours.tsv` for ours. This script joins by `(sheet, address)` and
reports mismatches on `text`, `width`, `height`, `numberFormat`,
`fontName`, `fontSize`, `bold`, `italic`, `wrapText`, `hAlign`, `vAlign`.

Geometry diffs use a 1.0-pt tolerance (matches `compare-vba-layout-
results.mjs`'s `coordThreshold`). Text comparison is exact (use the
report to spot Excel's `#####` overflow rule vs ours).
"""
import argparse
import csv
import sys


def load(path):
    rows = {}
    with open(path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f, delimiter='\t')
        for row in reader:
            key = (row.get('sheet', '').strip(), row.get('address', '').strip())
            rows[key] = row
    return rows


def fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('excel', help='Excel COM TSV (from extract-excel-layout.vbs)')
    parser.add_argument('ours', help='Our engine TSV (from spreadsheet-vba-layout-tsv.sh)')
    parser.add_argument('--tol', type=float, default=1.0,
                        help='Geometry tolerance in points (default 1.0)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Print every mismatch row')
    args = parser.parse_args()

    excel = load(args.excel)
    ours = load(args.ours)

    excel_only = set(excel) - set(ours)
    ours_only = set(ours) - set(excel)
    common = set(excel) & set(ours)

    print(f'excel: {len(excel)} rows; ours: {len(ours)} rows; common: {len(common)}')
    if excel_only:
        print(f'  excel-only addresses: {len(excel_only)} (e.g. {sorted(excel_only)[:5]})')
    if ours_only:
        print(f'  ours-only addresses: {len(ours_only)} (e.g. {sorted(ours_only)[:5]})')

    geom_fields = ['x', 'y', 'width', 'height']
    txt_fields = ['text', 'numberFormat', 'fontName', 'wrapText', 'hAlign', 'vAlign']
    num_fields = ['fontSize']
    bool_fields = ['bold', 'italic']

    mismatches = {f: [] for f in geom_fields + txt_fields + num_fields + bool_fields}

    for key in sorted(common):
        e, o = excel[key], ours[key]
        for f in geom_fields + num_fields:
            ev, ov = fnum(e.get(f)), fnum(o.get(f))
            if ev is None or ov is None:
                continue
            tol = args.tol if f in geom_fields else 0.1
            if abs(ev - ov) > tol:
                mismatches[f].append((key, ev, ov))
        for f in txt_fields:
            ev, ov = (e.get(f) or '').strip(), (o.get(f) or '').strip()
            if ev != ov:
                mismatches[f].append((key, ev, ov))
        for f in bool_fields:
            ev = (e.get(f) or '').strip().lower() in ('true', '1')
            ov = (o.get(f) or '').strip().lower() in ('true', '1')
            if ev != ov:
                mismatches[f].append((key, ev, ov))

    any_mismatch = False
    for f in geom_fields + num_fields + bool_fields + txt_fields:
        ms = mismatches[f]
        if not ms:
            continue
        any_mismatch = True
        print(f'\n=== {f}: {len(ms)} mismatches ===')
        sample = ms if args.verbose else ms[:10]
        for key, ev, ov in sample:
            sheet, addr = key
            print(f'  {sheet}!{addr}: excel={ev!r} ours={ov!r}')
        if not args.verbose and len(ms) > 10:
            print(f'  ... +{len(ms) - 10} more (use --verbose to see all)')

    if not any_mismatch:
        print('\nAll fields match within tolerance.')

    sys.exit(0 if not any_mismatch else 1)


if __name__ == '__main__':
    main()
