#!/usr/bin/env python3
"""Builds the printed full<->hash anchor / blind-holdout workbooks.

Purpose: measure Excel's *printed word advance* for formatted numbers in cells
wide enough not to clip (the Poppler word box), the right gap to the physical
cell edge (fill rectangle), and the tight physical cell width where Excel flips
from the full number to `###`.  Matched strings differ only in digit
composition (1s vs 8s) or length, and each has a literal-string twin; italic
and accounting-pad formats and several print scales are included.

Each test column carries its own solid fill (same palette/reader as
make-print-count-fixture.py) so `measure-print-fullhash-anchor.py` can read the
physical cell edges from the PDF.  Row 1 is the test cell; there is no row 2.

Split `anchor` (exploratory anchors, exported first) and `holdout` (blind:
predictions locked before this workbook is exported).  The holdout is a
face/format/scale combination absent from the FitCutoffEngine fixtures:
Cambria 16 upright 80 % `#,##0.00` and Trebuchet MS 12 italic 85 % `0.00%`.

    python3 tools/excel-parity/make-print-fullhash-anchor.py OUT.xlsx META.tsv anchor|holdout
"""
import importlib.util
import os
import sys
import zipfile
from xml.sax.saxutils import escape

_here = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location('countfixture', os.path.join(_here, 'make-print-count-fixture.py'))
countfixture = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(countfixture)

DECL = countfixture.DECL
MAIN = countfixture.MAIN
REL = countfixture.REL
PKG = countfixture.PKG
CT = countfixture.CT
fill_color = countfixture.fill_color
col_letter = countfixture.col_letter
attr = countfixture.attr

ACCT = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)'
# format code -> custom numFmtId
FMT_ID = {'0.00%': 166, '#,##0.00': 167, ACCT: 168}

WIDE = 22.0


def widths(start, stop, step=0.03125):
    n = int(round((stop - start) / step)) + 1
    return [round(start + i * step, 4) for i in range(n)]


def pct(v):
    return f'{v * 100:.2f}%'


def grp(v):
    return f'{v:,.2f}'


def acct(v):
    return f'${v:,.0f}'


# name, face, size, italic, scale, format, dense sweep, wide numeric values, literal twins, split
CASES = [
    ('A_V13up_pct', 'Verdana', 13, False, 100, '0.00%',
     widths(16.75, 17.75),
     [12345.6789, 11111.11, 88888.88, 1111.11, 8888.88, 111.11, 888.88, 111111.11],
     ['1234567.89%', '1111111.00%', '8888888.00%'], 'anchor'),
    ('A_V13it_pct', 'Verdana', 13, True, 100, '0.00%',
     widths(16.75, 18.0),
     [12345.6789, 11111.11, 88888.88, 1111.11, 8888.88], ['1234567.89%'], 'anchor'),
    ('A_V13up_pct75', 'Verdana', 13, False, 75, '0.00%',
     widths(16.5, 17.75),
     [12345.6789, 11111.11, 88888.88, 1111.11, 8888.88], ['1234567.89%'], 'anchor'),
    ('A_V13up_acct', 'Verdana', 13, False, 100, ACCT,
     widths(17.5, 19.0),
     [1234567, 1111111, 8888888, 111111, 888888], ['$1,234,567', '$1,111,111', '$8,888,888'], 'anchor'),
    ('A_G15it_grp', 'Georgia', 15, True, 90, '#,##0.00',
     widths(16.0, 17.25),
     [1234567, 1111111.11, 8888888.88, 111111.11, 888888.88, 11111.11, 88888.88, 1111.11, 8888.88,
      -1234567],
     ['1,234,567.00', '1,111,111.11', '8,888,888.88'], 'anchor'),
    ('A_G15up_grp', 'Georgia', 15, False, 90, '#,##0.00',
     widths(15.75, 17.0),
     [1234567, 1111111.11, 8888888.88, 111111.11, 888888.88], ['1,234,567.00', '1,111,111.11'], 'anchor'),
    ('A_G15it_pct', 'Georgia', 15, True, 100, '0.00%',
     widths(16.25, 17.5),
     [12345.6789, 11111.11, 88888.88, 1111.11, 8888.88], ['1234567.89%'], 'anchor'),
    # ---- blind holdout: face/format/scale absent from the engine fixtures ----
    ('H_Cambria16up_grp80', 'Cambria', 16, False, 80, '#,##0.00',
     widths(16.5, 17.75),
     [9876543.21, 1111111.11, 8888888.88, 111111.11, 888888.88], ['9,876,543.21', '1,111,111.11'], 'holdout'),
    ('H_Trebuchet12it_pct85', 'Trebuchet MS', 12, True, 85, '0.00%',
     widths(11.5, 12.75),
     [1234.5678, 1111.11, 8888.88, 111.11, 888.88], ['123456.78%', '111111.00%'], 'holdout'),
]

TEXT_OF = {'0.00%': pct, '#,##0.00': grp, ACCT: acct}


def build(out, meta_path, split):
    cases = [c for c in CASES if c[9] == split]
    if not cases:
        raise SystemExit(f'no cases for split {split}')

    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>']
    for _, face, size, italic, *_ in cases:
        fonts.append(f'{"<i/>" if italic else ""}<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>')

    max_cols = max(3 + len(c[6]) + len(c[7]) + len(c[8]) for c in cases)
    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    for i in range(max_cols):
        r, g, b = fill_color(i)
        fills.append(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{r:02X}{g:02X}{b:02X}"/>'
                     f'<bgColor indexed="64"/></patternFill></fill>')

    used_fmts = sorted({FMT_ID[c[5]] for c in cases})
    numfmts = ''.join(f'<numFmt numFmtId="{FMT_ID[code]}" formatCode="{attr(code)}"/>'
                      for code in FMT_ID if FMT_ID[code] in used_fmts)

    # Column 1 is a plain label column; test columns start at 2 (matching the count fixture).
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
           '<alignment horizontal="left"/></xf>']

    meta = ['sheet\tcolumn\trow\twidth_chars\talign\tface\tsize_pt\titalic\tscale\tkind\tformat\tvalue\trole\ttext']
    sheet_files = []
    for ci, (name, face, size, italic, scale, fmt, dense, wide_nums, wide_lits, _) in enumerate(cases):
        cols = []            # (width, kind, role, value_text)
        for v in wide_nums:
            cols.append((WIDE, 'wide', 'num', v))
        for lit in wide_lits:
            cols.append((WIDE, 'wide', 'lit', lit))
        for w in dense:
            cols.append((w, 'dense', 'num', wide_nums[0]))
        col_xml = ''.join(f'<col min="{2 + j}" max="{2 + j}" width="{w}" customWidth="1"/>'
                          for j, (w, _, _, _) in enumerate(cols))
        style_of = []
        control_of = []
        for j, (w, kind, role, v) in enumerate(cols):
            style_of.append(len(xfs))
            xfs.append(f'<xf numFmtId="{FMT_ID[fmt]}" fontId="{ci + 1}" fillId="{2 + j}" borderId="0" '
                       f'xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                       f'<alignment horizontal="right"/></xf>')
            # Row 2 is a non-overflow control of known text in the same face/scale/fill.
            control_of.append(len(xfs))
            xfs.append(f'<xf numFmtId="2" fontId="{ci + 1}" fillId="{2 + j}" borderId="0" '
                       f'xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1">'
                       f'<alignment horizontal="right"/></xf>')
        cells = [f'<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>{escape(name)}</t></is></c>']
        for j, (w, kind, role, v) in enumerate(cols):
            ref = f'{col_letter(2 + j)}1'
            if role == 'lit':
                cells.append(f'<c r="{ref}" s="{style_of[j]}" t="inlineStr"><is><t>{escape(v)}</t></is></c>')
                text = v
                value = ''
            else:
                cells.append(f'<c r="{ref}" s="{style_of[j]}"><v>{v}</v></c>')
                text = TEXT_OF[fmt](v)
                value = str(v)
            meta.append('\t'.join([name, str(2 + j), '1', f'{w:g}', 'right', face, str(size),
                                   '1' if italic else '0', str(scale), kind, fmt, value, role, text]))
        cells.append('</row>')
        cells.append(f'<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>control</t></is></c>')
        for j, (w, kind, role, v) in enumerate(cols):
            cells.append(f'<c r="{col_letter(2 + j)}2" s="{control_of[j]}"><v>-98.41</v></c>')
            meta.append('\t'.join([name, str(2 + j), '2', f'{w:g}', 'right', face, str(size),
                                   '1' if italic else '0', str(scale), 'control', '0.00', '-98.41',
                                   'num', '-98.41']))
        cells.append('</row>')
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if ci == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            f'<cols><col min="1" max="1" width="24" customWidth="1"/>{col_xml}</cols>'
            f'<sheetData>{"".join(cells)}</sheetData>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            f'<pageSetup paperSize="1" orientation="landscape" scale="{scale}"/></worksheet>')
        sheet_files.append((name, xml))

    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}">'
        f'<numFmts count="{len(used_fmts)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{f}</font>' for f in fonts) + '</fonts>'
        f'<fills count="{len(fills)}">' + ''.join(fills) + '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (name, _) in enumerate(sheet_files))
            + '</sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': DECL + (
            f'<Relationships xmlns="{PKG}">'
            + ''.join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>'
                      for i in range(len(sheet_files)))
            + f'<Relationship Id="rId{len(sheet_files) + 1}" Type="{REL}/styles" Target="styles.xml"/></Relationships>'),
        'xl/styles.xml': styles,
    }
    for i, (_, xml) in enumerate(sheet_files):
        files[f'xl/worksheets/sheet{i + 1}.xml'] = xml
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in files.items():
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, text)
    with open(meta_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(meta) + '\n')
    print(f'wrote {out} ({len(sheet_files)} sheets, '
          f'{sum(len(c[6]) + len(c[7]) + len(c[8]) for c in cases)} test columns) and {meta_path}')


if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[3] not in ('anchor', 'holdout'):
        raise SystemExit(__doc__)
    build(sys.argv[1], sys.argv[2], sys.argv[3])
