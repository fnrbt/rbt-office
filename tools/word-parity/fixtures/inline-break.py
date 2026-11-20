#!/usr/bin/env python3
"""Fixture family: a page or column break inside a paragraph's text (w:br with text before and after it).

    python3 tools/word-parity/fixtures/inline-break.py tests/word/fixtures/inline-break
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/inline-break --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/inline-break/word/*.pdf > tests/word/fixtures/inline-break/word/geometry.tsv

Hypothesis (docs/layout-parity-flow.md ledger): the text after the break starts the next page - in a section
with columns, after a column break, the next column; in one column a column break is a page break - at its
top, as a line of the same paragraph (no space before, no first-line indent); the line the break ends is as
high as its text. Each document: Calibri 11pt single-spaced (13.43pt), 12pt after each paragraph, a Letter
page with 1in margins; "Lead k" paragraphs, then the probe paragraph (first-line indent 0.5in, 12pt before):
"Probe start" <br> "Probe rest" (or several parts), then "Tail 1" ... Names:
  <kind>-c<cols>[-long][-m<mode>]   kind page / column / column2 (two column breaks) / page-col (a page break
                                    in a section with columns); cols 1 or 2; -long: each part three lines of text
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "inline-break"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
FILL = ("the quick brown fox jumps over the lazy dog while the five boxing wizards jump quickly "
        "and a wizard's job is to vex chumps quickly in fog ")

def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'

def document(kinds, cols, long, leads=4):
    body = [f'<w:p>{run(f"Lead {k}")}</w:p>' for k in range(1, leads + 1)]
    parts = ["Probe start", "Probe rest", "Probe third"][:len(kinds) + 1]
    runs = []
    for i, part in enumerate(parts):
        text = part + (" " + FILL * 3 if long else "")
        runs.append(run(text.strip()))
        if i < len(kinds):
            runs.append(f'<w:r><w:br w:type="{kinds[i]}"/></w:r>')
    body.append('<w:p><w:pPr><w:spacing w:before="240"/><w:ind w:firstLine="720"/></w:pPr>' + "".join(runs) + '</w:p>')
    body += [f'<w:p>{run(f"Tail {k}")}</w:p>' for k in range(1, 4)]
    c = f'<w:cols w:num="{cols}" w:space="720"/>' if cols > 1 else '<w:cols w:space="720"/>'
    body.append('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
                f'w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{c}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '</w:styles>')

def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')

def write(name, doc, mode=15):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", settings(mode))

count = 0
for name, kinds, cols in [("page", ["page"], 1), ("column", ["column"], 1), ("column", ["column"], 2),
                          ("column2", ["column", "column"], 2), ("page-col", ["page"], 2)]:
    for long in [False, True]:
        for mode in [15, 14]:
            write(f"{name}-c{cols}{'-long' if long else ''}{'-m14' if mode == 14 else ''}", document(kinds, cols, long), mode)
            count += 1
print(out, count)
