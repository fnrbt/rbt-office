#!/usr/bin/env python3
"""Fixture family: how tall Word makes the line of a paragraph that is an inline picture alone.

    python3 tools/word-parity/fixtures/picture-line.py [out dir]   (default tests/word/fixtures/picture-line)
    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/picline \\
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/picture-line --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/picture-line/word/*.pdf > tests/word/fixtures/picture-line/word/geometry.tsv

Word's picture line is the distance from the text of the line above it to the text of the line below it, less the
line above's own step. (Not ours yet - docs/layout-parity-flow.md ledger, open: the rule needs the picture's
wp:effectExtent, which the layout does not read; its check comes with it.)

From the showcase magazine-feature: a 104.5pt picture in a Georgia 10pt paragraph at 1.15 lines takes 1.5pt
more than the picture in Word, 2.7pt in ours. Every document: Letter, 1in margins, no space before or after;
"Above 1" ... "Above 3", the picture's paragraph (the picture alone, 2in wide), "Below 1" ... "Below 3", all
in the font, size and line spacing the name gives:
  <font>-<spacing>-h<height>   font cal11 (Calibri 11pt) or geo10 (Georgia 10pt); spacing s100 (single),
                               s115 (1.15 lines), s200 (double), l20 (at least 20pt), x20 (exactly 20pt);
                               the picture h50 / h100 / h104 (104.5pt) high
  m14-<...>                    some of them in compatibility mode 14 (the others 15)
  text-cal11-<spacing>-h50     a word after the picture in its line
  ea-m<mode>-h<height>         cal11-s115 with the text's East Asian font Malgun Gothic and language ko-KR (as
                               docx4j-repeating-section's defaults), h63 / h103 / h216 (216.78pt), mode 14 and 15
  after10-h<height>            ea-m14 with 10pt after every paragraph
No dependencies: the package is written as XML.
"""
import os, struct, sys, zipfile, zlib

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/picture-line")
os.makedirs(out, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
EMU = 12700


def png(rgb=(0x99, 0x99, 0x99), size=8):
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def inline(w, h):
    cx, cy = round(w * EMU), round(h * EMU)
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="0" name="grey.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


FONTS = {"cal11": ("Calibri", 22), "geo10": ("Georgia", 20)}
SPACING = {"s100": ("240", "auto"), "s115": ("276", "auto"), "s200": ("480", "auto"), "l20": ("400", "atLeast"), "x20": ("400", "exact")}


def write(name, font, spacing, height, mode=15, beside=False, east_asia=False, after=0):
    family, size = FONTS[font]
    line, rule = SPACING[spacing]
    ppr = f'<w:pPr><w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="{rule}"/></w:pPr>'
    text = lambda t: f'<w:p>{ppr}<w:r><w:t>{t}</w:t></w:r></w:p>'
    body = "".join(text(f"Above {k}") for k in (1, 2, 3)) + f'<w:p>{ppr}{inline(144, height)}{"<w:r><w:t>Beside</w:t></w:r>" if beside else ""}</w:p>' + "".join(text(f"Below {k}") for k in (1, 2, 3))
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              + (f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="Malgun Gothic" w:cs="{family}"/>'
                 f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
                 if east_asia else
                 f'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="{family}" w:cs="{family}"/>'
                 f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>')
              + '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    drels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
             '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/grey.png"/></Relationships>')
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings); z.writestr("word/media/grey.png", png())


for font in FONTS:
    for spacing in SPACING:
        for height in (50, 100):
            write(f"{font}-{spacing}-h{height}", font, spacing, height)
write("geo10-s115-h104", "geo10", "s115", 104.5)
for font, spacing, height in [("cal11", "s100", 100), ("cal11", "s115", 100), ("cal11", "s200", 50), ("cal11", "l20", 50), ("cal11", "x20", 50), ("geo10", "s115", 104.5)]:
    write(f"m14-{font}-{spacing}-h{int(height)}", font, spacing, height, mode=14)
# A picture's paragraph whose mark has an East Asian language and its theme's East Asian font (the docDefaults of
# docx4j-repeating-section: w:eastAsiaTheme minorEastAsia, w:lang w:eastAsia ko-KR - Malgun Gothic), mode 14 and 15.
for mode in (14, 15):
    for height in (63, 103.3, 216.78):
        write(f"ea-m{mode}-h{int(height)}", "cal11", "s115", height, mode=mode, east_asia=True)
# 10pt after every paragraph (as docx4j-repeating-section's).
for height in (63, 216.78):
    write(f"after10-h{int(height)}", "cal11", "s115", height, mode=14, east_asia=True, after=200)
for spacing in ("s100", "s115", "s200"):
    write(f"text-cal11-{spacing}-h50", "cal11", spacing, 50, beside=True)
print(out)
