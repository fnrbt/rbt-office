#!/usr/bin/env python3
"""Fixture family: where Word starts the endnotes at the end of the document.

    python3 tools/word-parity/fixtures/endnote-start.py <out dir>
    node tools/word-parity/page-parity.mjs --dist dist --dir <out dir> --refresh-word   (Word's PDFs in <out dir>/word)
    python3 tools/word-parity/pdf-geometry.py --paths <out dir>/word/*.pdf > <out dir>/word/geometry.tsv

Hypothesis (docs/layout-parity-flow.md ledger): the endnotes (w:endnotePr's default position, the end of the
document) start right below the last paragraph - its last line, not its space after - with endnotes.xml's
separator, a paragraph of its own as high as it is laid out (its spacing and line), then the notes; the
separator and the first note go on to the next page when they do not fit under the text. Each document:
n body paragraphs "Body k" of Calibri 11pt single-spaced (13.43pt), nothing before or after, on a Letter
page with 1in margins (648pt of text: 48 lines); an endnote reference in the last; the note one line of
"Note one" in EndnoteText (Calibri 10pt single, 12.21pt) unless named:
  base-n<n>       the separator as Word writes it (single-spaced, Calibri 11pt)
  sepdouble-n<n>  the separator double-spaced (w:line 480)
  sep20-n<n>      the separator's mark and run at 20pt
  sepbefore-n<n>  the separator 12pt before
  after12-n<n>    the last body paragraph 12pt after
  trail-n<n>      an empty paragraph after the one with the reference
  three-n<n>      the note three lines (two line breaks)
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else "endnote-start"
os.makedirs(out, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def document(n, after12=False, trail=False):
    paras = []
    for k in range(1, n + 1):
        last = k == n
        spacing = '<w:spacing w:after="240"/>' if last and after12 else ''
        ref = '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r>' if last else ''
        paras.append(f'<w:p><w:pPr>{spacing}</w:pPr><w:r><w:t>Body {k}</w:t></w:r>{ref}</w:p>')
    if trail:
        paras.append('<w:p/>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(paras)}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')

def endnotes(sep_ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', sep_rpr='', lines=1):
    text = '<w:r><w:t xml:space="preserve"> Note one</w:t></w:r>' + ''.join(f'<w:r><w:br/><w:t>Note line {k}</w:t></w:r>' for k in range(2, lines + 1))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes {W}>'
            f'<w:endnote w:type="separator" w:id="0"><w:p><w:pPr>{sep_ppr}{sep_rpr and "<w:rPr>" + sep_rpr + "</w:rPr>"}</w:pPr><w:r>{sep_rpr and "<w:rPr>" + sep_rpr + "</w:rPr>"}<w:separator/></w:r></w:p></w:endnote>'
            '<w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
            f'<w:endnote w:id="2"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>{text}</w:p></w:endnote>'
            '</w:endnotes>')

STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:basedOn w:val="DefaultParagraphFont"/>'
          '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="1"/></w:endnotePr>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')

def write(name, doc, notes):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
             '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '</Relationships>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS); z.writestr("word/endnotes.xml", notes)

count = 0
# Where the notes start (room to spare) and whether the separator and note fit at the foot:
# 46 x 13.43 + 13.43 + 12.21 = 643.3 of 648 fits, 47 does not.
for n in [10, 45, 46, 47]:
    write(f"base-n{n}", document(n), endnotes()); count += 1
for n in [10, 44, 45]:
    write(f"sepdouble-n{n}", document(n), endnotes(sep_ppr='<w:spacing w:after="0" w:line="480" w:lineRule="auto"/>')); count += 1
    write(f"sep20-n{n}", document(n), endnotes(sep_rpr='<w:sz w:val="40"/>')); count += 1
    write(f"sepbefore-n{n}", document(n), endnotes(sep_ppr='<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>')); count += 1
# The last paragraph's 12pt after: 46 lines and it fit only if it does not count.
for n in [10, 45, 46]:
    write(f"after12-n{n}", document(n, after12=True), endnotes()); count += 1
for n in [10, 45, 46]:
    write(f"trail-n{n}", document(n, trail=True), endnotes()); count += 1
# A three-line note (36.6pt): 44 lines leave room for it whole, 45 for two of its lines, 46 for one.
for n in [10, 44, 45, 46]:
    write(f"three-n{n}", document(n), endnotes(lines=3)); count += 1
print(out, count)
