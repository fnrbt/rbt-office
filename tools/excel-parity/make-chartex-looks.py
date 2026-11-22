#!/usr/bin/env python3
"""synthetic-chartex-looks.xlsx: Excel 2019's own chartEx charts (synthetic-chartex-zoo.xlsx, build-chartex-zoo.vbs)
with what COM cannot set written into their parts, for Excel to draw (each chart's picture through Chart.Export:
`BOOKFIRST=1 PNGDIR=dir probe-on-vm.sh chart-probe.vbs out.tsv synthetic-chartex-looks.xlsx`):

- "Hist1" / "Hist2": bins closed on the left (`intervalClosed="l"`).
- "HistCount": five bins (`cx:binCount`); "HistWidth": bins 150 wide from an underflow of 100 to an overflow of 800.
- "Treemap": labels of the categories, the parents' laid over (`parentLabelLayout` overlapping);
  "Treemap3": the parents' in a banner; "TreemapFlat": the categories' names and values.
- "Sunburst" / "Sunburst3": labels of the categories.
- "BoxTwo": the line through the means, no points inside the whiskers.

    python3 tools/excel-parity/make-chartex-looks.py
"""
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(os.path.dirname(os.path.dirname(HERE)), 'tests', 'test-documents-excel')
SOURCE = os.path.join(DOCS, 'synthetic-chartex-zoo.xlsx')
OUT = os.path.join(DOCS, 'synthetic-chartex-looks.xlsx')

LABELS = '<cx:dataLabels pos="{pos}"><cx:visibility seriesName="0" categoryName="1" value="{value}"/></cx:dataLabels>'


def labelled(xml, pos='inEnd', value='0'):
    return xml.replace('<cx:dataId val="0" />', LABELS.format(pos=pos, value=value) + '<cx:dataId val="0" />', 1)


EDITS = {
    'xl/charts/chartEx1.xml': lambda x: x.replace('<cx:binning intervalClosed="r" />', '<cx:binning intervalClosed="l" />'),
    'xl/charts/chartEx2.xml': lambda x: x.replace('<cx:binning intervalClosed="r" />', '<cx:binning intervalClosed="l" />'),
    'xl/charts/chartEx9.xml': lambda x: x.replace('<cx:binning intervalClosed="r" />', '<cx:binning intervalClosed="r"><cx:binCount val="5"/></cx:binning>'),
    'xl/charts/chartEx10.xml': lambda x: x.replace('<cx:binning intervalClosed="r" />', '<cx:binning intervalClosed="r" underflow="100" overflow="800"><cx:binSize val="150"/></cx:binning>'),
    'xl/charts/chartEx17.xml': lambda x: labelled(x).replace('<cx:layoutPr />', '<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>'),
    'xl/charts/chartEx21.xml': lambda x: labelled(x).replace('<cx:layoutPr />', '<cx:layoutPr><cx:parentLabelLayout val="banner"/></cx:layoutPr>'),
    'xl/charts/chartEx19.xml': lambda x: labelled(x, value='1'),
    'xl/charts/chartEx18.xml': lambda x: labelled(x, pos='ctr'),
    'xl/charts/chartEx22.xml': lambda x: labelled(x, pos='ctr'),
    'xl/charts/chartEx15.xml': lambda x: x.replace('<cx:statistics quartileMethod="exclusive" /></cx:layoutPr>',
                                                   '<cx:visibility meanLine="1" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive" /></cx:layoutPr>'),
}


def main():
    src = zipfile.ZipFile(SOURCE)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in EDITS:
                text = data.decode('utf-8')
                edited = EDITS[item.filename](text)
                assert edited != text, item.filename
                data = edited.encode('utf-8')
            info = zipfile.ZipInfo(item.filename, date_time=(2026, 11, 22, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            out.writestr(info, data)
    print(OUT)


if __name__ == '__main__':
    main()
