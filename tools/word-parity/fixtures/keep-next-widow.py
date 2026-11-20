#!/usr/bin/env python3
"""Fixture family: a keepNext paragraph (or a chain of two) at a page's foot, before a paragraph that widow
control, keepLines or its own length keeps from leaving a single line behind.

    python3 tools/word-parity/fixtures/keep-next-widow.py tests/word/fixtures/keep-next-widow
    node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/keep-next-widow --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/keep-next-widow/word/*.pdf > tests/word/fixtures/keep-next-widow/word/geometry.tsv

docx4j-repeating-section's "RepeatingSection positives" (Heading 1, keepNext) and the keepNext paragraph under it go to
page 7 in Word, where the 2-line paragraph after them would have had one line on page 6. Each document: Letter, 1in
margins, Calibri 11pt single-spaced, no space after anywhere (48 lines to a page, 13.43pt each); filler lines "A k",
then "H-probe" (keepNext), in the -hk documents "K-probe" (keepNext) too, then "P1" ... "Pn" (one paragraph, its lines
w:br-broken), then "C 1". Name: <chain>-n<n>-r<r>[-nowc|-kl][-m14]
  chain  h (H-probe alone) or hk (H-probe, K-probe)
  n      the lines of the paragraph P
  r      how many of P's lines fit on page 1 after the chain
  -nowc  widow control off (w:widowControl w:val="0" in Normal), -kl  P has keepLines
  -m14   compatibility mode 14 (else 15)
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "keep-next-widow"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr>')
LINES = 48


def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def document(chain, n, r, keep_lines):
    head = ["H-probe"] + (["K-probe"] if chain == "hk" else [])
    fillers = LINES - len(head) - r
    body = [para(f"A {k}") for k in range(1, fillers + 1)]
    body += [para(h, '<w:keepNext/>') for h in head]
    runs = '<w:r><w:br/></w:r>'.join(f'<w:r><w:t xml:space="preserve">P{k} line</w:t></w:r>' for k in range(1, n + 1))
    body.append(f'<w:p><w:pPr>{"<w:keepLines/>" if keep_lines else ""}</w:pPr>{runs}</w:p>')
    body.append(para("C 1"))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{SECT}</w:body></w:document>'


def styles(widow):
    wc = '' if widow else '<w:pPr><w:widowControl w:val="0"/></w:pPr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>{wc}</w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
            '</w:styles>')


def settings(mode):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


def write(name, doc, mode, widow=True):
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
        z.writestr("word/styles.xml", styles(widow)); z.writestr("word/settings.xml", settings(mode))


count = 0
for chain in ["h", "hk"]:
    for n in [2, 3, 4, 5]:
        for r in [1, 2, 3]:
            write(f"{chain}-n{n}-r{r}", document(chain, n, r, False), 15); count += 1
            if chain == "hk":
                write(f"{chain}-n{n}-r{r}-m14", document(chain, n, r, False), 14); count += 1
for n in [2, 3, 4]:
    for r in [1, 2]:
        write(f"hk-n{n}-r{r}-nowc", document("hk", n, r, False), 15, widow=False); count += 1
for r in [2, 3]:
    write(f"hk-n4-r{r}-kl", document("hk", 4, r, True), 15); count += 1
print(out, count)
