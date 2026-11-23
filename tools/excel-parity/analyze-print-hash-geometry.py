#!/usr/bin/env python3
"""Summarises measure-print-hash-geometry.py's rows: edge slopes per fixed hash count, count transitions, italic
translations, and prediction accuracy.

    python3 tools/excel-parity/analyze-print-hash-geometry.py MEASUREMENTS.tsv OUT_SUMMARY.tsv
"""
import collections
import statistics
import sys


def lin(xs, ys):
    n = len(xs)
    if n < 2:
        return None
    mx, my = statistics.mean(xs), statistics.mean(ys)
    den = sum((x - mx) ** 2 for x in xs)
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / den if den else None


def main():
    src, out_path = sys.argv[1:3]
    rows = [l.rstrip('\n').split('\t') for l in open(src, encoding='utf-8')]
    head = rows[0]
    rows = [dict(zip(head, r)) for r in rows[1:]]
    out = ['kind\tsheet\talign\tface\tsize_pt\titalic\tscale\tn\tsamples\t'
           'slope_xMinLeft\tslope_xMaxLeft\tslope_rightGap\trun_advance\ttransition_from\t'
           'transition_to\tdW\tdxMinLeft\tdxMaxLeft\trun_advance_after\tnotes']
    groups = collections.defaultdict(list)
    by_sheet = collections.defaultdict(dict)
    for r in rows:
        n = int(r['nHashes'])
        # literal runs longer than 8 are two adjacent cells' text merged by the PDF word extractor (spill)
        if n > 0 and not (r['kind'] == 'literal' and n > 8):
            groups[(r['kind'], r['sheet'], r['align'], r['face'], r['size_pt'], r['italic'], r['scale'], n)].append(r)
        if r['row'] == '1' and r['kind'] == 'overflow':
            by_sheet[r['sheet']][float(r['width_chars'])] = r
    for k, g in sorted(groups.items()):
        W = [float(r['cellWidth']) for r in g]
        a = lin(W, [float(r['xMinLeft']) for r in g])
        b = lin(W, [float(r['xMaxLeft']) for r in g])
        c = lin(W, [float(r['rightGap']) for r in g])
        adv = statistics.mean(float(r['runWidth']) / int(r['nHashes']) for r in g)
        out.append('\t'.join([k[0], k[1], k[2], k[3], k[4], k[5], k[6], str(k[7]), str(len(g)),
                              f'{a:+.4f}' if a is not None else '', f'{b:+.4f}' if b is not None else '',
                              f'{c:+.4f}' if c is not None else '', f'{adv:.4f}', '', '', '', '', '', '', '']))
    for sheet, items in sorted(by_sheet.items()):
        items = sorted(items.items())
        for (w1, r1), (w2, r2) in zip(items, items[1:]):
            n1, n2 = int(r1['nHashes']), int(r2['nHashes'])
            if n1 and n2 and n1 != n2:
                out.append('\t'.join(['transition', sheet, r1['align'], r1['face'], r1['size_pt'], r1['italic'],
                                      r1['scale'], f'{n1}->{n2}', '', '', '', '', '',
                                      str(w1), str(w2), f'{float(r2["cellWidth"]) - float(r1["cellWidth"]):+.3f}',
                                      f'{float(r2["xMinLeft"]) - float(r1["xMinLeft"]):+.3f}',
                                      f'{float(r2["xMaxLeft"]) - float(r1["xMaxLeft"]):+.3f}',
                                      f'{float(r2["runWidth"]) / n2:.4f}', 'centred: dxMin=-a/2+dW/2, dxMax=+a/2+dW/2']))
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'wrote {len(out) - 1} summary rows to {out_path}')


if __name__ == '__main__':
    main()
