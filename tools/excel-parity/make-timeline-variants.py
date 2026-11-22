#!/usr/bin/env python3
"""Writes copies of tests/test-documents-excel/synthetic-slicers.xlsx whose timeline selects other periods at other levels,
for how Excel 2019 words a timeline's selection (its selection label): open each in Excel (shoot-sheets.ps1, sheet
"Report") and read the label above the periods. (SlicerCaches.Add2(..., xlTimeline) is refused from a script, so the
timelines cannot be made in Excel by VBS.)

    python3 tools/excel-parity/make-timeline-variants.py OUTDIR

Writes OUTDIR/synthetic-timeline-<case>.xlsx for each case below: level (0 years, 1 quarters, 2 months, 3 days), first
and last day selected.
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.normpath(os.path.join(HERE, '..', '..', 'tests', 'test-documents-excel', 'synthetic-slicers.xlsx'))

CASES = {
    'months-one-year': (2, '2012-03-01', '2012-12-31'),
    'months-two-years': (2, '2012-11-01', '2013-02-28'),
    'month': (2, '2012-06-01', '2012-06-30'),
    'quarters-one-year': (1, '2012-04-01', '2012-09-30'),
    'quarters-two-years': (1, '2012-10-01', '2013-03-31'),
    'quarter': (1, '2012-04-01', '2012-06-30'),
    'days-one-month': (3, '2012-03-05', '2012-03-20'),
    'days-two-months': (3, '2012-03-28', '2012-04-03'),
    'days-two-years': (3, '2012-12-30', '2013-01-02'),
    'day': (3, '2012-03-05', '2012-03-05'),
    'years': (0, '2012-01-01', '2013-12-31'),
    'year': (0, '2012-01-01', '2012-12-31'),
}


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    for name, (level, first, last) in CASES.items():
        source = zipfile.ZipFile(SOURCE)
        target = zipfile.ZipFile(os.path.join(out_dir, 'synthetic-timeline-%s.xlsx' % name), 'w', zipfile.ZIP_DEFLATED)
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == 'xl/timelines/timeline1.xml':
                text = data.decode('utf-8')
                text = re.sub(r' level="\d"', ' level="%d"' % level, text)
                text = re.sub(r' selectionLevel="\d"', ' selectionLevel="%d"' % level, text)
                data = text.encode('utf-8')
            elif item.filename == 'xl/timelineCaches/timelineCache1.xml':
                text = data.decode('utf-8')
                text = re.sub(r'<selection startDate="[^"]*" endDate="[^"]*"/>',
                              '<selection startDate="%sT00:00:00" endDate="%sT00:00:00"/>' % (first, last), text)
                data = text.encode('utf-8')
            target.writestr(item, data)
        target.close()
        source.close()


if __name__ == '__main__':
    main()
