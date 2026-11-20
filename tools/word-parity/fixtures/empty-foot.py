#!/usr/bin/env python3
"""Fixture family: an empty paragraph at a page's foot - does it go on to the next page as a line would?

    python3 tools/word-parity/fixtures/empty-foot.py tests/word/fixtures/empty-foot
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/empty-foot --refresh-word

docx4j-getting-started p21: Word's page starts with the empty Consolas 8pt paragraph after the property
lines that end p20 (its first text 9.36pt below the margin); ours had the page start at the text. Each
document: A4 (697.95pt of text between 1in margins), Consolas 8pt single-spaced lines (9.37pt; 74 fit),
nothing before or after; n lines "Line<k>", then the probe, then "Nextq" and "Lastq".
Names: e<n> - after n lines an empty paragraph (its mark Consolas 8pt); t<n> - a paragraph of one tab
(Consolas 12pt, as docx4j-getting-started's code lines); x<n> - a text line "Probeq" (the control).
Word's page 2 first line (page-lines.vbs) says where the probe went. letter<k>-<probe>74: the same after k
Letter pages (a section of its own, a next-page break into the A4 one), as docx4j-getting-started's p20.
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CON = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'


def para(content, size=16):
    return (f'<w:p><w:pPr><w:rPr>{CON}<w:sz w:val="16"/></w:rPr></w:pPr>'
            + (f'<w:r><w:rPr>{CON}<w:sz w:val="{size}"/></w:rPr>{content}</w:r>' if content else '') + '</w:p>')


def write(out, name, n, probe, letter_pages=0):
    letter = ''
    if letter_pages:
        letter = ''.join(para(f'<w:t>Letter{k}</w:t>') + para('<w:br w:type="page"/>') for k in range(1, letter_pages))
        letter += (para(f'<w:t>Letter{letter_pages}</w:t>')[:-6] .replace('</w:rPr></w:pPr>', '</w:rPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
                   '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr>', 1) + '</w:p>')
    body = letter + ''.join(para(f'<w:t>Line{k}</w:t>') for k in range(1, n + 1))
    body += {"e": para(""), "t": para('<w:tab/>', 24), "x": para('<w:t>Probeq</w:t>')}[probe]
    body += para('<w:t>Nextq</w:t>') + para('<w:t>Lastq</w:t>')
    doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
           '<w:sectPr><w:pgSz w:w="11907" w:h="16839" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
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


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "empty-foot"
    os.makedirs(out, exist_ok=True)
    for probe in "etx":
        for n in (72, 73, 74):
            write(out, f"{probe}{n}", n, probe)
    for k in (1, 3, 19):
        for probe in "ex":
            write(out, f"letter{k}-{probe}74", 74, probe, k)
