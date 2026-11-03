#!/usr/bin/env python3
"""Prints the built-in table styles of `presetTableStyles.xml` (ECMA-376 Part 1, in
OfficeOpenXML-SpreadsheetMLStyles.zip; Apache POI ships the same file as
poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml) as the `presets`
table of src/Spreadsheet/TableStyles.fs. The file numbers the dxfs of a style from 1.

    python3 tools/excel-parity/make-table-style-presets.py presetTableStyles.xml
"""
import xml.etree.ElementTree as ET, sys, re
src = sys.argv[1]
r = ET.parse(src).getroot()
KINDS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe',
         'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell']
STY = {'thin': '1', 'medium': '2', 'thick': '3', 'double': 'd', 'hair': 'h', 'dotted': 'o', 'dashed': 'a'}
EDGE = {'top': 't', 'bottom': 'b', 'left': 'l', 'right': 'r', 'horizontal': 'h', 'vertical': 'v'}


def tint(t):
    if t is None:
        return ''
    v = float(t)
    s = ('%.6f' % v).rstrip('0').rstrip('.')
    return ',' + s


def col(c):
    if c is None:
        return None
    if 'theme' in c.attrib:
        return c.attrib['theme'] + tint(c.attrib.get('tint'))
    if 'rgb' in c.attrib:
        return '#' + c.attrib['rgb'][-6:]
    raise Exception(c.attrib)


def dxf(d):
    out = []
    f = d.find('fill/patternFill')
    if f is not None:
        c = f.find('bgColor') if f.find('bgColor') is not None else f.find('fgColor')
        out.append('F' + col(c))
    fo = d.find('font')
    if fo is not None:
        if fo.find('b') is not None:
            out.append('B')
        if fo.find('i') is not None:
            out.append('I')
        c = fo.find('color')
        if c is not None:
            out.append('C' + col(c))
    b = d.find('border')
    if b is not None:
        for e in b:
            if e.tag in EDGE and 'style' in e.attrib:
                out.append(EDGE[e.tag] + STY[e.attrib['style']] + ':' + col(e.find('color')))
    return ';'.join(out)


lines = []
for st in r:
    name = st.tag
    if not name.startswith('TableStyle'):
        continue
    dx = st.findall('dxfs/dxf')
    ts = st.find('tableStyles/tableStyle')
    els = []
    for e in ts.findall('tableStyleElement'):
        k = e.attrib['type']
        if k not in KINDS:
            continue
        size = e.attrib.get('size')
        els.append(k + ('*' + size if size and size != '1' else '') + '=' + dxf(dx[int(e.attrib['dxfId']) - 1]))
    lines.append((name, '|'.join(els)))


def key(n):
    m = re.match(r'TableStyle(Light|Medium|Dark)(\d+)', n[0])
    return ({'Light': 0, 'Medium': 1, 'Dark': 2}[m.group(1)], int(m.group(2)))


lines.sort(key=key)
total = sum(len(a) + len(b) for a, b in lines)
print('// %d styles, %d chars' % (len(lines), total), file=sys.stderr)
for n, s in lines:
    print('      "%s", "%s"' % (n, s))
