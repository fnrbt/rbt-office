#!/usr/bin/env python3
"""Fixture family: what of the next paragraph has to fit for a keep-with-next paragraph to stay at a page's foot?

    python3 tools/word-parity/fixtures/keep-next-fit.py tests/word/fixtures/keep-next-fit
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/keep-next-fit --refresh-word

docx4j-getting-started p39: its Heading 1 "SmartArt" (keepNext, 24pt before, 12pt after) at the foot, the next
paragraph one line: Word starts p40 with the heading; ours kept it - the group counted the heading's lines and
the next one's first line, not the heading's space after between them. Each document: Letter, 1in margins
(648pt), Calibri 11pt single-spaced lines (13.43pt), nothing before or after but where named; n lines "Line<k>",
the keep-with-next paragraph "Keepq", then the next paragraph "Nextq ..." and "Lastq".
Names: a<after>-b<before>-n<n>[-w][-s2]: Keepq's space after, Nextq's space before (in points). With n lines
above, Keepq's text ends at 13.43 (n + 1) and Nextq's first line's at 13.43 (n + 2) + the larger of the two
spaces: 46 lines and 12pt end Nextq's at 656.6 (over), 45 at 643.2; with no spaces 46 lines end it at 644.6,
47 at 658.0 (over). -w: Nextq three lines long (widow control: its first line alone may not stay); -s2: after
a section of one page (a next-page break), laid out section by section.
Word's page 2 first line (page-lines.vbs) says whether Keepq went on.
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LONG = "Nextq " + "and some more words to fill the line " * 9


def para(text, before=0, after=0, keep=False, sect=""):
    ppr = (f'<w:pPr>{"<w:keepNext/>" if keep else ""}<w:spacing w:before="{before * 20}" w:after="{after * 20}" w:line="240" w:lineRule="auto"/>{sect}</w:pPr>')
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def write(out, name, after, before, n, widow=False, s2=False):
    body = para("Firstq", sect=SECT) if s2 else ""
    body += ''.join(para(f"Line{k}") for k in range(1, n + 1))
    body += para("Keepq", after=after, keep=True)
    body += para(LONG if widow else "Nextq", before=before) + para("Lastq")
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}{SECT}</w:body></w:document>'
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             f'<Relationship Id="rId1" Type="{R}styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels); z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings)


def variants():
    for s2 in (False, True):
        sfx = "-s2" if s2 else ""
        for after, before, n in ((12, 0, 45), (12, 0, 46), (0, 12, 45), (0, 12, 46), (12, 6, 46), (0, 0, 46), (0, 0, 47)):
            yield f"a{after}-b{before}-n{n}{sfx}", after, before, n, False, s2
        for n in (45, 46):
            yield f"a0-b0-n{n}-w{sfx}", 0, 0, n, True, s2


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "keep-next-fit"
    os.makedirs(out, exist_ok=True)
    for name, *rest in variants():
        write(out, name, *rest)
