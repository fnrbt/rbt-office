#!/usr/bin/env python3
"""Fixture family: which tab stops a paragraph has when its style has some and it states its own.

    python3 tools/word-parity/fixtures/tab-merge.py [out dir]   (default tests/word/fixtures/tab-merge)
    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/tabmerge \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/tab-merge --refresh-word
    for f in tests/word/fixtures/tab-merge/word/*.pdf; do n=$(basename $f .pdf); pdftotext -bbox $f - | grep '<word' |
      sed -E "s/.*xMin=\"([0-9.]+)\" yMin=\"([0-9.]+)\" xMax=\"([0-9.]+)\".*>(.*)<\/word>/$n\t\4\t\1\t\3\t\2/"; done \
      > tests/word/fixtures/tab-merge/word/words.tsv      (fixture, word, left, right, top in points)

tests/word/tab-merge.fsx checks ours against Word's: where each word after a tab starts and ends.

From the showcase magazine-feature: its running head (style Header: a centre stop at 4680, a right one at 9360;
the paragraph's own right stop at 9411) has "Feature  ·  Issue 15" centred at 4680 in Word - ours took the
paragraph's stops only, set it right at 9411 and wrapped "15". Hypothesis: a paragraph's stops are its style's
(down its basedOn chain) with its own added, a w:tab w:val="clear" taking away the style's at that position.
Every document: Letter, 1in margins, Calibri 11pt, one paragraph "Aaa<tab>Bbb<tab>Ccc<tab>Ddd" in style Tabbed
(basedOn Normal); stops in twips from the left margin:
  style-only      Tabbed: centre 4680, right 9360; the paragraph none
  para-only       Tabbed none; the paragraph: right 9411
  merge           Tabbed: centre 4680, right 9360; the paragraph: right 9411
  merge-left      Tabbed: centre 4680, right 9360; the paragraph: left 1440
  clear           Tabbed: centre 4680, right 9360; the paragraph: clear 4680
  clear-add       Tabbed: centre 4680, right 9360; the paragraph: clear 4680, left 2880
  same-pos        Tabbed: centre 4680; the paragraph: right 4680 (the paragraph's kind at one position)
  chain           Base: left 1440; Tabbed basedOn Base: centre 4680; the paragraph: right 9000
  chain-clear     Base: left 1440, centre 4680; Tabbed basedOn Base: clear 1440; the paragraph none
No dependencies: the package is written as XML.
"""
import os, sys, zipfile

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/tab-merge")
os.makedirs(out, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def tabs(stops):
    if not stops: return ""
    return "<w:tabs>" + "".join(f'<w:tab w:val="{kind}" w:pos="{pos}"/>' for kind, pos in stops) + "</w:tabs>"


def write(name, style_stops, para_stops, base_stops=None):
    based = '<w:basedOn w:val="Base"/>' if base_stops is not None else '<w:basedOn w:val="Normal"/>'
    base = ('<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr>{tabs(base_stops)}</w:pPr></w:style>') if base_stops is not None else ""
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
              f'{base}<w:style w:type="paragraph" w:styleId="Tabbed"><w:name w:val="Tabbed"/>{based}<w:pPr>{tabs(style_stops)}</w:pPr></w:style></w:styles>')
    run = lambda t: f'<w:r><w:t>{t}</w:t></w:r>'
    tab = '<w:r><w:tab/></w:r>'
    body = (f'<w:p><w:pPr><w:pStyle w:val="Tabbed"/>{tabs(para_stops)}</w:pPr>{run("Aaa")}{tab}{run("Bbb")}{tab}{run("Ccc")}{tab}{run("Ddd")}</w:p>'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}</w:body></w:document>'
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:defaultTabStop w:val="720"/>'
                '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings)


HEADER = [("center", 4680), ("right", 9360)]
write("style-only", HEADER, [])
write("para-only", [], [("right", 9411)])
write("merge", HEADER, [("right", 9411)])
write("merge-left", HEADER, [("left", 1440)])
write("clear", HEADER, [("clear", 4680)])
write("clear-add", HEADER, [("clear", 4680), ("left", 2880)])
write("same-pos", [("center", 4680)], [("right", 4680)])
write("chain", [("center", 4680)], [("right", 9000)], base_stops=[("left", 1440)])
write("chain-clear", [("clear", 1440)], [], base_stops=[("left", 1440), ("center", 4680)])
print(out)
