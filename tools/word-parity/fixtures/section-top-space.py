#!/usr/bin/env python3
"""Fixture family: the space before of the first paragraph of a section that starts a new page, by what
the section before it is.

    python3 tools/word-parity/fixtures/section-top-space.py tests/word/fixtures/section-top-space
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/section-top-space --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/section-top-space/word/*.pdf > tests/word/fixtures/section-top-space/word/geometry.tsv

docx4j-getting-started's "Creating a new docx" (Heading 1, 24pt before) starts the document's last section
(next page, A4) after a continuous section of two columns (Letter): Word sets it at the top margin, ours 24pt
lower. page-top-space's `sect` (after a plain next-page break) keeps the space. Each document: Calibri 11pt
single-spaced, nothing before or after but on the probe "P-probe" (12pt before: its top 74.16 at the margin,
86.18 with the space); 1in margins; section A "A k" (6 lines, Letter, one column), then:
  base          the probe's section (next page, Letter)
  cont          a continuous section B "B k" (6 lines, one column) before it
  cols          B in two columns
  a4            the probe's section A4 (A Letter)
  cols-a4       B in two columns and the probe's section A4 (as docx4j-getting-started)
  cols-a4-h1    the same, the probe keepNext and keepLines (as its Heading 1)
  after<A>-b<B> A's last paragraph A twips after (auto: w:afterAutospacing), the probe B twips before
in compatibility mode 14 (-m14) and 15.
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "section-top-space"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
LETTER = '<w:pgSz w:w="12240" w:h="15840"/>'
A4 = '<w:pgSz w:w="11907" w:h="16839" w:code="9"/>'
MAR = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'

def sect(size, cols=1, kind=None):
    t = f'<w:type w:val="{kind}"/>' if kind else ''
    c = f'<w:cols w:num="{cols}" w:space="708"/>' if cols > 1 else '<w:cols w:space="708"/>'
    return f'<w:sectPr>{t}{size}{MAR}{c}</w:sectPr>'

def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'

def document(b, probe_size, keep, after=None, before=240):
    body = [para(f"A {k}") for k in range(1, 6)]
    if after is not None:
        spacing = '<w:spacing w:after="100" w:afterAutospacing="1"/>' if after == "auto" else f'<w:spacing w:after="{after}"/>'
        body.append(para("A 6", spacing + sect(LETTER)))
    elif b is None:
        body.append(para("A 6", sect(LETTER)))
    else:
        body.append(para("A 6", sect(LETTER)))
        body += [para(f"B {k}") for k in range(1, 6)]
        body.append(para("B 6", sect(LETTER, b, "continuous")))
    keeps = '<w:keepNext/><w:keepLines/>' if keep else ''
    body.append(para("P-probe", keeps + f'<w:spacing w:before="{before}"/>'))
    body += [para(f"C {k}") for k in range(1, 4)]
    body.append(sect(probe_size))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}</w:body></w:document>'

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '</w:styles>')

def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')

def write(name, doc, mode):
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
for name, b, size, keep in [("base", None, LETTER, False), ("cont", 1, LETTER, False), ("cols", 2, LETTER, False),
                            ("a4", None, A4, False), ("cols-a4", 2, A4, False), ("cols-a4-h1", 2, A4, True)]:
    for mode in [14, 15]:
        write(name + ("-m14" if mode == 14 else ""), document(b, size, keep), mode); count += 1
# The last paragraph of the section before with space after (twips; auto: w:afterAutospacing, 14pt), the probe
# 12 or 24pt before.
for after in [120, 240, 360, "auto"]:
    for before in [240, 480]:
        for mode in [14, 15]:
            write(f"after{after}-b{before}" + ("-m14" if mode == 14 else ""), document(None, LETTER, False, after, before), mode); count += 1
print(out, count)
