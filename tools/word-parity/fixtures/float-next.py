#!/usr/bin/env python3
"""Fixture family: the paragraphs after a square-wrapped picture's own - do their lines go round it, in one
column and in a section with columns, and where does its box (distB under it) end for them?

    python3 tools/word-parity/fixtures/float-next.py [out dir]   (default tests/word/fixtures/float-next)

then, to ask Word (winvm) and compare with ours (Word's pages and PDFs in <out dir>/word):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/float-next --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/float-next/word/*.pdf > tests/word/fixtures/float-next/word/geometry.tsv

tests/word/float-next.fsx checks ours against geometry.tsv (each line's page, top and left edge).

Hypothesis (docs/layout-parity-flow.md ledger): a picture placed from its paragraph keeps its box - distB
under it - clear of the text of the paragraphs after its own too, in a section with columns as on a page:
a later paragraph's line goes beside it when its top (below the paragraph's space before, the space after
above it) is above the picture's bottom plus distB, whatever the line's height.
The showcase's magazine-feature.docx (two columns) has its turbines picture 8.5pt down "The engineers talk
about capacity factors ..." with 9pt distB: in Word the next paragraph's first line, 1.46pt into that
distance, is beside it; ours set it full width.

Every document: Letter, 1in margins, Calibri 11pt single-spaced (13.43pt lines), nothing before or after,
justified. A ("a01 ... a20": one line, or two in a column), B (the anchor's, one short line "b01 ... b05"), C (one
short line), D (a long paragraph "d001 ..."). The picture: a grey PNG 1in wide at the margin's (column's) left, B's top,
square, both sides, its anchor at B's start; its distances from text 0 but distB. D's first line's top is
2 lines (26.86pt) below B's top. Names:

  <cols>-d<d>-<m|p>05      distB d pt, the picture's bottom plus d 0.5pt above (m) / below (p) D's first
                           line's top
  <cols>-a6-d9-<m|p>05     C 6pt after: D's first line 6pt lower, the picture's bottom with it
  <cols>-b6-d9-<m|p>05     D 6pt before, alike
  <cols>-b6-d9-m<55|65>    D 6pt before, the picture's bottom plus 9pt 0.5pt below / above D's paragraph's
                           top (its space before's)
  <cols>-l15-d9-<m|p>05    D at 1.5 lines (its lines 20.15pt): its first line's top where it was
  <cols>-d9-p05-c14        the same in Word 2010's layout (compatibilityMode 14)
  <cols>: one (one column) or col (two columns 0.5in apart, all of it in the first)
No dependencies: the package is written as XML, the PNG with zlib.
"""
import os, struct, sys, zipfile, zlib

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/float-next")
os.makedirs(out, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
EMU = 12700          # per point
LINE = 13.4307       # Word's single line of Calibri 11pt
IN = 72.0


def png(rgb=(0x99, 0x99, 0x99), size=8):
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def words(prefix, count, digits):
    return " ".join(f"{prefix}{k:0{digits}d}" for k in range(1, count + 1))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0, after=0, line=240):
    sp = f'<w:spacing w:before="{before * 20}" w:after="{after * 20}" w:line="{line}" w:lineRule="auto"/>'
    return f'<w:p><w:pPr>{sp}<w:jc w:val="both"/></w:pPr>{content}</w:p>'


def picture(h, dist_b, rel_h):
    w = IN
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="{round(dist_b * EMU)}" distL="0" distR="0" simplePos="0" '
            'relativeHeight="251658241" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{rel_h}"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{round(w * EMU)}" cy="{round(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Float 1"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="0" name="grey.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:anchor></w:drawing></w:r>')


def document(body, columns):
    cols = f'<w:cols w:num="{columns}" w:space="720"/>' if columns > 1 else ""
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{cols}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' + SINGLE + '</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>')


def write(name, body, columns, compat=15):
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
                f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document(body, columns))
        z.writestr("word/styles.xml", STYLES); z.writestr("word/settings.xml", settings); z.writestr("word/media/grey.png", png())


A = para(run(words("a", 20, 2)))
C_WORDS = run(words("c", 5, 2))
D_WORDS = run(words("d", 200, 3))


def body(dist_b, step, rel_h, c_after=0, d_before=0, d_line=240):
    # D's first line's top below B's top: B's line and C's, C's space after or D's space before (the larger).
    room = 2 * LINE + max(c_after, d_before)
    h = room - dist_b + step
    return (A + para(picture(h, dist_b, rel_h) + run(words("b", 5, 2)))
            + para(C_WORDS, after=c_after) + para(D_WORDS, before=d_before, line=d_line))


for cols, columns, rel_h in (("one", 1, "margin"), ("col", 2, "column")):
    for tag, step in (("m05", -0.5), ("p05", 0.5)):
        for d in (0, 9, 18):
            write(f"{cols}-d{d}-{tag}", body(d, step, rel_h), columns)
        write(f"{cols}-a6-d9-{tag}", body(9, step, rel_h, c_after=6), columns)
        write(f"{cols}-b6-d9-{tag}", body(9, step, rel_h, d_before=6), columns)
        write(f"{cols}-l15-d9-{tag}", body(9, step, rel_h, d_line=360), columns)
    # (D's space before: its paragraph's top 6pt above its first line's - the picture's bottom plus distB
    # 0.5pt below / above that.)
    for tag, step in (("m55", -5.5), ("m65", -6.5)):
        write(f"{cols}-b6-d9-{tag}", body(9, step, rel_h, d_before=6), columns)
    write(f"{cols}-d9-p05-c14", body(9, 0.5, rel_h), columns, compat=14)
    write(f"{cols}-d9-m05-c14", body(9, -0.5, rel_h), columns, compat=14)
print(out)
