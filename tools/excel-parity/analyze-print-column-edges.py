#!/usr/bin/env python3
"""Summarises measure-print-column-edges.py's vector measurements.

Per sheet and width character: border-centreline width (true printed cell
width) against the stored-width layout conversion and GeometryModel's
scaled-dot prediction; fill expansion and adjacent-fill overlap; and the
full/hash decision against each width model.  Also fits the repeated
equal-width block per page, separating edge i from edge 0, page breaks, page
origins and margins.

    python3 tools/excel-parity/analyze-print-column-edges.py META PRED PRED_GEOM RESULTS OUT.tsv
"""
import collections
import statistics
import sys


def load(path):
    rows = [l.rstrip('\n').split('\t') for l in open(path, encoding='utf-8') if not l.startswith('#')]
    return [dict(zip(rows[0], r)) for r in rows[1:]]


def main():
    pred_path, geom_path, results_path, out_path = sys.argv[2:6]
    pred = {(r['sheet'], r['column']): r for r in load(pred_path)}
    geom = {(r['sheet'], r['column']): r for r in load(geom_path)}
    rows = load(results_path)

    sheets = sorted({r['sheet'] for r in rows})
    out = ['sheet\twidth_chars\tn\tborder_mean\tborder_min\tborder_max\tfill_mean\t'
           'fill_minus_border\tfill_overlap\tlayout_W\tgeom_W\tborder_minus_layout\tborder_minus_geom\t'
           'decision\tdecision_match_fill\tdecision_match_layout\tdecision_match_geom']
    tally = {s: collections.Counter() for s in sheets}
    for sheet in sheets:
        srows = [r for r in rows if r['sheet'] == sheet]
        for w in sorted({r['width_chars'] for r in srows}, key=float):
            rs = [r for r in srows if r['width_chars'] == w]
            pick = rs[0]
            bw = [float(r['border_width']) for r in rs]
            fw = [float(r['fill_width']) for r in rs]
            fb = [float(r['fill_minus_border']) for r in rs]
            ov = [float(r['fill_overlap_next']) for r in rs if r['fill_overlap_next']]
            key = (sheet, pick['column'])
            layout = float(pred[key]['pred_W_layout_pt'])
            g = float(geom[key]['pred_border_width_pt'])
            decision = pick['decision']
            m_fill = decision == pick['pred_fill_decision']
            m_layout = decision == pred[key]['pred_layout_decision']
            m_geom = decision == geom[key]['pred_decision']
            tally[sheet]['n'] += len(rs)
            tally[sheet]['fill'] += len(rs) if m_fill else 0
            tally[sheet]['layout'] += len(rs) if m_layout else 0
            tally[sheet]['geom'] += len(rs) if m_geom else 0
            out.append('\t'.join([
                sheet, w, str(len(rs)), f'{statistics.mean(bw):.4f}', f'{min(bw):.4f}', f'{max(bw):.4f}',
                f'{statistics.mean(fw):.4f}', f'{statistics.mean(fb):.4f}',
                f'{statistics.mean(ov):.4f}' if ov else '', f'{layout:.3f}', f'{g:.3f}',
                f'{statistics.mean(bw) - layout:+.4f}', f'{statistics.mean(bw) - g:+.4f}', decision,
                'yes' if m_fill else 'no', 'yes' if m_layout else 'no', 'yes' if m_geom else 'no']))

    print('Decision agreement (fill width / layout width / GeometryModel width):')
    for s in sheets:
        t = tally[s]
        print(f'  {s:10s} {t["fill"]}/{t["n"]}  {t["layout"]}/{t["n"]}  {t["geom"]}/{t["n"]}')

    print('\nFive columns of interest at 90 % (widths 16.2188-16.3438):')
    for r in rows:
        if (r['sheet'] in ('E90_edge', 'H90_shift') and r['block'] == 'vary'
                and 16.2188 <= float(r['width_chars']) <= 16.3438):
            key = (r['sheet'], r['column'])
            print(f'  {r["sheet"]:9s} col{r["column"]:>3s} w={r["width_chars"]:>8s} '
                  f'border={r["border_width"]:>7s} fill={r["fill_width"]:>7s} '
                  f'layout={pred[key]["pred_W_layout_pt"]} geom={geom[key]["pred_border_width_pt"]} '
                  f'decision={r["decision"]}')

    print('\nFill/stroke expansion (all columns):')
    fb = [float(r['fill_minus_border']) for r in rows]
    ov = [float(r['fill_overlap_next']) for r in rows if r['fill_overlap_next']]
    print(f'  fill_width - border_width: mean {statistics.mean(fb):.4f} '
          f'range {min(fb):.3f}..{max(fb):.3f}  n={len(fb)}')
    print(f'  adjacent fill overlap:     mean {statistics.mean(ov):.4f} '
          f'range {min(ov):.3f}..{max(ov):.3f}  n={len(ov)}')

    print('\nRepeated equal-width block, page-local edge series:')
    for sheet in sheets:
        pages = collections.defaultdict(list)
        for r in rows:
            if r['sheet'] == sheet and r['block'] == 'repeat':
                pages[int(r['page'])].append(r)
        pitches, residuals, origins = [], [], []
        for p, rs in sorted(pages.items()):
            rs.sort(key=lambda r: int(r['column']))
            bounds = [float(rs[0]['border_left'])] + [float(r['border_right']) for r in rs]
            n = len(rs)
            xs = list(range(n + 1))
            mx, my = sum(xs) / len(xs), sum(bounds) / len(bounds)
            slope = (sum((x - mx) * (y - my) for x, y in zip(xs, bounds))
                     / sum((x - mx) ** 2 for x in xs))
            pitches.append(slope)
            residuals += [y - (bounds[0] + slope * i) for i, y in enumerate(bounds)]
            origins.append((p, n, round(bounds[0], 2)))
        print(f'  {sheet:10s} pages={len(pages)} pitch mean {statistics.mean(pitches):.4f} '
              f'range {min(pitches):.4f}..{max(pitches):.4f}  '
              f'max |edge_i - (edge_0 + i*pitch)| = {max(abs(v) for v in residuals):.4f}  '
              f'first-edge per page={[o[2] for o in origins]}')
    with open(out_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    print(f'\nwrote {len(out) - 1} summary rows to {out_path}')


if __name__ == '__main__':
    main()
