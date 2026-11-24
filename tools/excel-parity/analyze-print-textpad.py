#!/usr/bin/env python3
"""Check the locked text-pad holdout predictions against a measured comparison.

    python3 tools/excel-parity/analyze-print-textpad.py \
        tools/excel-parity/print-textpad-holdout-predictions.tsv \
        COMPARISON.tsv WORKBOOK > RESULTS.tsv

COMPARISON.tsv is measure-print-textpad.py's long table for one holdout workbook. The checks are the
relations registered in print-textpad-holdout-predictions.tsv before the workbook's Excel export; values
are taken from the same PDF's F00 row so page geometry cancels. `excel_ok` verifies the prediction on
Excel itself; `ours_ok` verifies that the candidate reproduces Excel's relation within 0.10 pt.
"""
import collections
import sys


def load(path):
    header = open(path, encoding='utf-8').readline().rstrip('\n').split('\t')
    rows = []
    for line in open(path, encoding='utf-8'):
        if line.startswith('#') or not line.strip() or line.startswith('sheet\t'):
            continue
        rows.append(dict(zip(header, line.rstrip('\n').split('\t'))))
    return rows


def main():
    predictions, comparison, workbook = sys.argv[1:4]
    old_comparison = sys.argv[4] if len(sys.argv) > 4 else None
    rows = load(comparison)
    old_rows = load(old_comparison) if old_comparison else []
    old_table = collections.defaultdict(dict)
    for row in old_rows:
        old_table[(row['sheet'], int(row['row']), row['kind'], row['format_name'])][row['side']] = row
    table = collections.defaultdict(dict)
    for row in rows:
        key = (row['sheet'], int(row['row']), row['kind'], row['format_name'])
        table[key][row['side']] = row
    sheets = sorted({key[0] for key in table})
    out = ['workbook\tsheet\ttest\texcel_value\tours_value\texpected\texcel_ok\tours_ok']

    def cell(sheet, row, kind, side):
        names = [k for k in table if k[0] == sheet and k[1] == row and k[2] == kind]
        if not names:
            return None
        return table[names[0]].get(side)

    def find(sheet, prefix, kind='text', side='excel'):
        for key in sorted(table, key=lambda k: k[1]):
            if key[0] == sheet and key[2] == kind and key[3].startswith(prefix):
                return table[key].get(side)
        return None

    def old_find(sheet, prefix, kind='num', side='ours'):
        for key in sorted(old_table, key=lambda k: k[1]):
            if key[0] == sheet and key[2] == kind and key[3].startswith(prefix):
                return old_table[key].get(side)
        return None

    def num(record, field):
        return float(record[field]) if record is not None else None

    def add(sheet, test, excel_value, ours_value, expected, excel_ok, ours_ok):
        out.append('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s' % (
            workbook, sheet, test,
            '' if excel_value is None else f'{excel_value:.3f}',
            '' if ours_value is None else f'{ours_value:.3f}',
            expected, 'yes' if excel_ok else 'no', 'yes' if ours_ok else 'no'))

    for sheet in sheets:
        align = sheet.rsplit('_', 1)[-1]
        base_e, base_o = find(sheet, 'F00'), find(sheet, 'F00', side='ours')
        f02e, f02o = find(sheet, 'F02'), find(sheet, 'F02', side='ours')
        f03e, f03o = find(sheet, 'F03'), find(sheet, 'F03', side='ours')
        f04e, f04o = find(sheet, 'F04'), find(sheet, 'F04', side='ours')
        f05e, f05o = find(sheet, 'F05'), find(sheet, 'F05', side='ours')

        if align == 'right':
            rel2_e = num(f02e, 'gap_right') - num(base_e, 'gap_right')
            rel3_e = num(f03e, 'gap_right') - num(base_e, 'gap_right')
            rel2_o = num(f02o, 'gap_right') - num(base_o, 'gap_right')
            rel3_o = num(f03o, 'gap_right') - num(base_o, 'gap_right')
            add(sheet, 'REL_F02_R', rel2_e, rel2_o, '0 < d <= 7.0', 0 < rel2_e <= 7.0,
                abs(rel2_o - rel2_e) <= 0.10)
            add(sheet, 'REL_F03_R', rel3_e, rel3_o, 'd >= REL_F02_R + 1.0',
                rel3_e >= rel2_e + 1.0, abs(rel3_o - rel3_e) <= 0.10)
        if align == 'left':
            rel4_e = num(f04e, 'xMin') - num(base_e, 'xMin')
            rel4_o = num(f04o, 'xMin') - num(base_o, 'xMin')
            add(sheet, 'REL_F04_L', rel4_e, rel4_o, '0 < d <= 7.0', 0 < rel4_e <= 7.0,
                abs(rel4_o - rel4_e) <= 0.10)
        if align == 'center':
            d2_e = num(f02e, 'xMin') - num(base_e, 'xMin')
            d2_o = num(f02o, 'xMin') - num(base_o, 'xMin')
            d3_e = num(f03e, 'xMin') - num(base_e, 'xMin')
            d3_o = num(f03o, 'xMin') - num(base_o, 'xMin')
            d4_e = num(f04e, 'xMin') - num(base_e, 'xMin')
            d4_o = num(f04o, 'xMin') - num(base_o, 'xMin')
            rel2 = num(f02e, 'gap_right') - num(base_e, 'gap_right')
            rel3 = num(f03e, 'gap_right') - num(base_e, 'gap_right')
            rel4 = num(f04e, 'xMin') - num(base_e, 'xMin')
            f05_e = num(f05e, 'xMin') - num(base_e, 'xMin')
            f05_o = num(f05o, 'xMin') - num(base_o, 'xMin')
            add(sheet, 'CENTER_F02_LEFT', d2_e, d2_o, 'negative half advance',
                d2_e < 0 and abs(d4_e + d2_e) <= 0.05, abs(d2_o - d2_e) <= 0.10)
            add(sheet, 'CENTER_F03_WIDER', d3_e, d3_o, 'more negative than F02',
                d3_e < d2_e, abs(d3_o - d3_e) <= 0.10)
            add(sheet, 'CENTER_F04_RIGHT', d4_e, d4_o, 'positive half advance',
                d4_e > 0 and abs(d4_e + d2_e) <= 0.05, abs(d4_o - d4_e) <= 0.10)
            add(sheet, 'CENTER_F05_ZERO', f05_e, f05_o, '0',
                abs(f05_e) <= 0.05, abs(f05_o - f05_e) <= 0.10)

        for name in ('F02', 'F03', 'F04', 'F05'):
            e = find(sheet, name)
            o = find(sheet, name, side='ours')
            add(sheet, f'{name}_TEXT', num(e, 'xMax'), num(o, 'xMax'),
                e['text'] if e else '?', e is not None and e['text'] == 'Ab1',
                o is not None and o['text'] == 'Ab1' and abs(num(o, 'xMax') - num(e, 'xMax')) <= 0.65)

        for name, expected, test in (
                ('F06', '(Ab1', 'F06_TEXT'), ('F07', '_Ab1', 'F07_TEXT'),
                ('F08', '_Ab1', 'F08_TEXT'), ('F09', 'Ab1', 'F09_EQ_F05'),
                ('F10', 'Pre Ab1', 'F10_TEXT'), ('C01', 'Ab1', 'C01_C02_C05'),
                ('C02', 'Ab1', 'C01_C02_C05'), ('C03', None, 'C03_C04_HIDE'),
                ('C04', None, 'C03_C04_HIDE')):
            e = find(sheet, name, 'text')
            o = find(sheet, name, 'text', side='ours')
            ok_text = (e is None and expected is None) or (e is not None and e['text'] == expected)
            ok_ours = ((e is None and o is None) or
                       (e is not None and o is not None and o['text'] == expected
                        and abs(num(o, 'xMin') - num(e, 'xMin')) <= 0.65
                        and abs(num(o, 'xMax') - num(e, 'xMax')) <= 0.65))
            add(sheet, test + '_' + name, num(e, 'xMin') if e else None,
                num(o, 'xMin') if o else None, expected or 'hidden', ok_text, ok_ours)

        for name in ('F02', 'F03', 'F04', 'F05', 'F06', 'F07', 'F08', 'F10'):
            e = find(sheet, name, 'num')
            o = find(sheet, name, 'num', side='ours')
            prev = old_find(sheet, name, 'num', 'ours')
            add(sheet, f'NUM_UNCH_{name}', num(prev, 'xMin') if prev else None,
                num(o, 'xMin') if o else None, 'old == new; both match Excel',
                prev is not None and o is not None and abs(num(o, 'xMin') - num(prev, 'xMin')) <= 0.001,
                abs(num(o, 'xMin') - num(e, 'xMin')) <= 0.70)

    print('\n'.join(out))


if __name__ == '__main__':
    main()
