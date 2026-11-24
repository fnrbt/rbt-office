#!/usr/bin/env python3
"""Writes synthetic workbooks for where Excel 2019's PDF sets a text value under a custom number format's
fourth (text) section (docs/excel-parity-tracker.md, "Printed text under the fourth number-format section").

Each alignment sheet has one row per format code. Column B holds the text value "Ab1" with a yellow fill;
column C holds the numeric control 1234.5 with a pale-blue fill, under the same format. Both fills are solid
rectangles whose printed edges are the physical cell edges, so the measurement can compare a word box with
its cell without depending on page geometry.

Sets:
  train   - Aptos Narrow 14 at 100%: @, four-section @, @_), @_W, leading/both/repeated `_` pads, quoted
            and escaped underscores, standard accounting, quoted prefix/suffix, two- and three-section @
            controls, empty fourth, all-empty, escaped suffix.
  holdout - the same text-section cases on Georgia 12 at 75% and Tahoma 10 at 100% (independent face and
            print scale).
  escapes - numeric formats with OOXML `\\X` escapes, to check that preserving escapes for the text section
            does not change numeric rendering.

    python3 tools/excel-parity/make-print-textpad-fixture.py OUT.xlsx CELLS.tsv SET [--face NAME] [--size N]
        [--scale N] [--prefix P]

Deterministic: run it again and the bytes are the same.
"""
import argparse
import zipfile
from xml.sax.saxutils import escape

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml'

TRAIN_FORMATS = [
    ('F00_plain_at', '@'),
    ('F01_four_at', '0;[Red]-0;0;@'),
    ('F02_four_at_pad_r', '0;[Red]-0;0;@_)'),
    ('F03_four_at_pad_W', '0;[Red]-0;0;@_W'),
    ('F04_four_lead_pad', '0;[Red]-0;0;_(@'),
    ('F05_four_both_pad', '0;[Red]-0;0;_(@_)'),
    ('F06_four_repeat_pad', '0;[Red]-0;0;__(@__'),
    ('F07_quoted_underscore', '0;[Red]-0;0;"_"@'),
    ('F08_escaped_underscore', '0;[Red]-0;0;\\_@'),
    ('F09_accounting', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'),
    ('F10_literal_prefix', '0;[Red]-0;0;"Pre "@'),
    ('F11_at_pad_r_dup', '0;[Red]-0;0;@_)'),
    ('C01_two_sections_at', '0;@'),
    ('C02_three_sections_at', '0;0;@'),
    ('C03_four_empty', '0;0;0;'),
    ('C04_all_empty', ';;;'),
    ('C05_three_empty_third', '0;0;'),
    ('C06_lead_pad_suffix', '0;[Red]-0;0;_(@" units"'),
    ('C07_quoted_brackets', '0;[Red]-0;0;"["@"]"'),
    ('C08_escaped_underscore_suffix', '0;[Red]-0;0;@\\_'),
]
HOLDOUT_FORMATS = [
    ('F00_plain_at', '@'),
    ('F02_four_at_pad_r', '0;[Red]-0;0;@_)'),
    ('F03_four_at_pad_W', '0;[Red]-0;0;@_W'),
    ('F04_four_lead_pad', '0;[Red]-0;0;_(@'),
    ('F05_four_both_pad', '0;[Red]-0;0;_(@_)'),
    ('F06_four_repeat_pad', '0;[Red]-0;0;__(@__'),
    ('F07_quoted_underscore', '0;[Red]-0;0;"_"@'),
    ('F08_escaped_underscore', '0;[Red]-0;0;\\_@'),
    ('F09_accounting', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'),
    ('F10_literal_prefix', '0;[Red]-0;0;"Pre "@'),
    ('C01_two_sections_at', '0;@'),
    ('C02_three_sections_at', '0;0;@'),
    ('C03_four_empty', '0;0;0;'),
    ('C04_all_empty', ';;;'),
]
ESCAPES_FORMATS = [
    ('E00_plain', '0.00'),
    ('E01_esc_underscore_prefix', '\\_0.00'),
    ('E02_esc_underscore_suffix', '0.00\\_'),
    ('E03_esc_at_prefix', '\\@0.00'),
    ('E04_esc_at_suffix', '0.00\\@'),
    ('E05_esc_hyphen_suffix', '0.00\\-'),
    ('E06_date_esc_hyphens', 'yyyy\\-mm\\-dd'),
    ('E07_esc_underscore_grouped', '\\_#,##0.00_)'),
    ('E08_quoted_underscore_prefix', '"_"0.00'),
    ('E09_esc_backslash', '0.00\\\\'),
    ('E10_esc_space', '0.00\\ '),
]
ALIGNS = ['left', 'center', 'right']
TEXT = 'Ab1'
NUM = 1234.5
FILL_TEXT = 'FFFFF2CC'
FILL_NUM = 'FFDDEBF7'


def attr(text):
    return escape(text, {'"': '&quot;'})


def build(out, cells_path, formats, face, size, scale, prefix):
    fonts = ['<sz val="11"/><name val="Calibri"/><family val="2"/>',
             f'<sz val="{size}"/><name val="{attr(face)}"/><family val="2"/>']
    fills = ('<fill><patternFill patternType="none"/></fill>'
             '<fill><patternFill patternType="gray125"/></fill>'
             f'<fill><patternFill patternType="solid"><fgColor rgb="{FILL_TEXT}"/><bgColor indexed="64"/></patternFill></fill>'
             f'<fill><patternFill patternType="solid"><fgColor rgb="{FILL_NUM}"/><bgColor indexed="64"/></patternFill></fill>')
    borders = ('<border><left/><right/><top/><bottom/><diagonal/></border>'
               '<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right>'
               '<top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom>'
               '<diagonal/></border>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    style = {}
    for ai, aname in enumerate(ALIGNS):
        for fi in range(len(formats)):
            for kind, fill in (('text', 2), ('num', 3)):
                style[(ai, fi, kind)] = len(xfs)
                xfs.append(
                    f'<xf numFmtId="{200 + fi}" fontId="1" fillId="{fill}" borderId="1" xfId="0" '
                    f'applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">'
                    f'<alignment horizontal="{aname}"/></xf>')
    numfmts = ''.join(f'<numFmt numFmtId="{200 + i}" formatCode="{attr(code)}"/>'
                      for i, (_, code) in enumerate(formats))
    styles = DECL + (
        f'<styleSheet xmlns="{MAIN}"><numFmts count="{len(formats)}">{numfmts}</numFmts>'
        f'<fonts count="{len(fonts)}">' + ''.join(f'<font>{x}</font>' for x in fonts) + '</fonts>'
        f'<fills count="4">{fills}</fills><borders count="2">{borders}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">' + ''.join(xfs) + '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    sheet_files = []
    manifest = []
    for ai, aname in enumerate(ALIGNS):
        rows = []
        for fi, (fname, code) in enumerate(formats):
            r = fi + 1
            st = style[(ai, fi, 'text')]
            sn = style[(ai, fi, 'num')]
            rows.append(
                f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>{fname}</t></is></c>'
                f'<c r="B{r}" s="{st}" t="inlineStr"><is><t>{escape(TEXT)}</t></is></c>'
                f'<c r="C{r}" s="{sn}"><v>{NUM}</v></c></row>')
            manifest.append((f'{prefix}{aname}', ai + 1, r, 'text', fname, code, TEXT))
            manifest.append((f'{prefix}{aname}', ai + 1, r, 'num', fname, code, str(NUM)))
        xml = DECL + (
            f'<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
            f'<sheetViews><sheetView{" tabSelected=\"1\"" if ai == 0 else ""} workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            '<cols><col min="1" max="1" width="30" customWidth="1"/>'
            '<col min="2" max="3" width="20" customWidth="1"/></cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            f'<pageSetup paperSize="1" orientation="portrait" scale="{scale}"/></worksheet>')
        sheet_files.append((f'{prefix}{aname}', xml))
    over = ''.join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="{CT}.worksheet+xml"/>'
                   for i in range(len(sheet_files)))
    files = {
        '[Content_Types].xml': DECL + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>{over}</Types>'),
        '_rels/.rels': DECL + (
            f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
        'xl/workbook.xml': DECL + (
            f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets>'
            + ''.join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
                      for i, (n, _) in enumerate(sheet_files)) + '</sheets></workbook>'),
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
    with open(cells_path, 'w', encoding='utf-8') as f:
        f.write('sheet\tindex\trow\tkind\tformat_name\tformat\tvalue\n')
        for row in manifest:
            f.write('\t'.join(map(str, row)) + '\n')
    print(out)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('out')
    parser.add_argument('cells')
    parser.add_argument('set', choices=['train', 'holdout', 'escapes'])
    parser.add_argument('--face', default='Aptos Narrow')
    parser.add_argument('--size', type=int, default=14)
    parser.add_argument('--scale', type=int, default=100)
    parser.add_argument('--prefix', default=None)
    args = parser.parse_args()
    formats = {'train': TRAIN_FORMATS, 'holdout': HOLDOUT_FORMATS, 'escapes': ESCAPES_FORMATS}[args.set]
    prefix = args.prefix if args.prefix is not None else ('T_' if args.set == 'train' else ('E_' if args.set == 'escapes' else 'H_'))
    build(args.out, args.cells, formats, args.face, args.size, args.scale, prefix)


if __name__ == '__main__':
    main()
