#!/usr/bin/env python3
"""Fixture family: how Word wraps text round a floating object (square and top-and-bottom wrapping).

    python3 tools/word-parity/fixtures/float-wrap.py [out dir]   (default tests/word/fixtures/float-wrap)

then, to ask Word (winvm) and compare with ours (Word's pages and PDFs in <out dir>/word):

    WORD_VM_DIR=C:/Users/olivia/RbtExcelParity/<own dir> \
      node tools/word-parity/page-parity.mjs --dist dist --dir tests/word/fixtures/float-wrap --refresh-word
    python3 tools/word-parity/pdf-geometry.py tests/word/fixtures/float-wrap/word/*.pdf > tests/word/fixtures/float-wrap/word/geometry.tsv

tests/word/float-wrap.fsx checks ours against geometry.tsv (each line's fragments: page, top, left,
right, first and last word; each picture's box).

Hypotheses (docs/layout-parity-flow.md ledger):
- a floating object positioned from its paragraph stands where that paragraph starts plus its offset,
  wherever in the paragraph its anchor run is, and every line of that paragraph (and after it) whose
  box overlaps the object's box widened by distT / distB goes beside it (square) or below it, distB
  under it (top and bottom);
- beside a square-wrapped object the text goes on both sides (bothSides), only on its left or right
  (left / right), or on the wider side (largest), a fragment no narrower than some least width;
- in a section with columns `column` is the anchor's column, `margin` the page's margins.

Every document: Letter, 1in margins, Calibri 11pt single-spaced (13.43pt lines), nothing before or
after, justified paragraphs of numbered words: A (2 lines, "a01 a02 ..."), B (the anchor's, ~14
lines, "b001 ..."), C ("c01 ..."). The picture is a plain grey PNG; its distances from text 0 unless
the name says. Only what the name says varies:

  sq-start / sq-mid / sq-end    a 2in x 1in picture at the margin's left, the paragraph's top; square,
                                both sides, 9pt all round; its anchor run at B's start / middle / end
  sq-end-heading, sq-start-heading   the same after a heading (14pt bold, 12pt before, 4pt after, keepNext)
                                as the showcase's diagrams.docx has it
  sq-end-keepnext               the same with A keeping with B
  sq-h4<m|p><f>                 its bottom f/10 pt above (m) or below (p) the 4th line's bottom: which
                                lines does it take?
  sq-top-<m|p><f>               1in high, its top 3 lines down B less / plus f/10 pt: which lines above it
                                stay full width?
  sq-distb18, sq-distt18        4 lines high, 18pt from text below / above (offset 4 lines for distT)
  side-<both|left|right|largest>-<c|l|r>   1.5in x 1in, 9pt round, centred (c) or 1in (l) / 3.5in (r)
                                from the margin, with that wrapText
  gap-<g>, gapr-<g>             2in x 1in, both sides, g pt from the margin's left (right) edge to it
  tb-start, tb-end              2in x 1in at the margin's left, B's top, top and bottom
  tb-off<k>                     its top k/10 lines down B
  tb-dist                       top 2.5 lines down B, 9pt above and 18pt below it
  tb-shape                      a shape (wps, 5.4in x 1in, 0.2cm down B) as the magazine's pull quote
  tb-low                        B starting 3 lines above the bottom of the page, a 2in-high picture
  col-<ref>                     two columns (0.5in apart); B at the top of the second (a column break
                                before it); a 1.5in x 1in picture at B's top, square, 9pt round,
                                posH: column right / column 36pt / margin 36pt / margin right / page 324pt
  col-cross                     the picture anchored in column 1 (A), 300pt from the margin (into column 2)
  col-tb-low                    the magazine's case: a shape topped-and-bottomed in a paragraph near
                                the first column's bottom (col-tb-low-start: its anchor at B's start)
  colq-k<k>-<end|start>-c<14|15>   the pull quote (216 x 73.7pt, 0.2cm down, top and bottom) at the end /
                                start of a 6-line B, k one-line paragraphs above it in the first of two columns;
                                -full: the second column full of text; -a6 / -a4: B 6pt / 4pt after;
                                colq-k46-a6-noshape: no shape
  mag-*                         the showcase's magazine-feature.docx, its quote moved / resized, Word 2019's
                                layout, without hyphenation, indent, space after ... (see Batch 7)
  cbr-*                         a column break's paragraph mark: text ending in one, a paragraph of only one
                                (12pt after, its mark 20pt, its break 20pt), two in a row
  <name>-c14                    the same in Word 2010's layout (compatibilityMode 14)
  gapil-<g>[-c14|-c12|-nocompat], gapilr-<g>   5pt words ("il") in a band of g pt (gapil-165t: 16.5pt),
                                in that compatibility mode
  split2-<sq|tb>-<start|end>    B's first line 4 lines above the page's bottom (A: 44 one-line paragraphs)
  kl-<one|two|none|tb>-c<14|15> a keepLines paragraph of one / two lines (none: not kept; tb: top and
                                bottom) between A and B, the picture 9pt round
  rep, rep-<c15|nokeeplines|single|noempty>    the showcase's diagrams.docx "On site" (Segoe UI 10pt at
                                1.15 lines, a Georgia keepLines heading, Word 2010's layout) and one thing
                                changed
No dependencies: the package is written as XML, the PNG with zlib.
"""
import os, struct, sys, zipfile, zlib

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/word/fixtures/float-wrap")
os.makedirs(out, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="wps"')
EMU = 12700          # per point
LINE = 13.4307       # Word's single line of Calibri 11pt
SINGLE = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'


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


def para(content, ppr=""):
    return f'<w:p><w:pPr>{ppr}{SINGLE}<w:jc w:val="both"/></w:pPr>{content}</w:p>'


def pos(tag, rel, val):
    body = f"<wp:align>{val}</wp:align>" if isinstance(val, str) else f"<wp:posOffset>{round(val * EMU)}</wp:posOffset>"
    return f'<wp:position{tag} relativeFrom="{rel}">{body}</wp:position{tag}>'


_ids = [0]


def anchor(graphic, w, h, ph, pv, wrap="square", side="bothSides", dist=(0, 0, 0, 0)):
    """w, h in pt; ph / pv: (relativeFrom, pt offset or alignment); dist: (top, bottom, left, right) pt."""
    _ids[0] += 1
    i = _ids[0]
    t, b, l, r = (round(d * EMU) for d in dist)
    wrap_xml = f'<wp:wrapSquare wrapText="{side}"/>' if wrap == "square" else "<wp:wrapTopAndBottom/>"
    return (f'<w:r><w:drawing><wp:anchor distT="{t}" distB="{b}" distL="{l}" distR="{r}" simplePos="0" relativeHeight="{251658240 + i}" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>{pos("H", *ph)}{pos("V", *pv)}'
            f'<wp:extent cx="{round(w * EMU)}" cy="{round(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_xml}'
            f'<wp:docPr id="{i}" name="Float {i}"/><wp:cNvGraphicFramePr/>{graphic}</wp:anchor></w:drawing></w:r>')


def picture(w, h, ph, pv, **kw):
    graphic = ('<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
               '<pic:nvPicPr><pic:cNvPr id="0" name="grey.png"/><pic:cNvPicPr/></pic:nvPicPr>'
               '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'
               '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    return anchor(graphic, w, h, ph, pv, **kw)


def shape(w, h, ph, pv, text="Pull quote", **kw):
    graphic = ('<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
               f'<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'
               '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
               f'<wps:txbx><w:txbxContent><w:p><w:pPr>{SINGLE}</w:pPr>{run(text)}</w:p></w:txbxContent></wps:txbx>'
               '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
               '</wps:wsp></a:graphicData></a:graphic>')
    return anchor(graphic, w, h, ph, pv, **kw)


def b_para(obj, where="start", count=200, ppr=""):
    text = words("b", count, 3)
    if where == "start":
        content = obj + run(text)
    elif where == "end":
        content = run(text) + obj
    else:
        half = text.split(" ")
        content = run(" ".join(half[:count // 2]) + " ") + obj + run(" ".join(half[count // 2:]))
    return para(content, ppr)


# (A is 2 lines however its text is measured: Word fits 25 of these words in a justified line in
# compatibilityMode 15, 24 in 14; lines of 40 words break the same in both.)
A = para(run(words("a", 40, 2)))
# Many one-line paragraphs: n lines of text wherever a line breaks.
def lines_of(n, prefix="a"):
    return "".join(para(run(f"{prefix}{k:03d}")) for k in range(1, n + 1))
C = para(run(words("c", 60, 2)))
HEADING = ('<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80" w:line="240" w:lineRule="auto"/></w:pPr>'
           '<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Heading</w:t></w:r></w:p>')


def document(body, columns=None):
    cols = f'<w:cols w:num="{columns}" w:space="720"/>' if columns else ""
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{cols}</w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' + SINGLE + '</w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>')


def write(name, body, columns=None, compat=15, styles=None):
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
    compat_xml = (f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat>'
                  if compat else "")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'{compat_xml}</w:settings>')
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct); z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", drels); z.writestr("word/document.xml", document(body, columns))
        z.writestr("word/styles.xml", styles or STYLES); z.writestr("word/settings.xml", settings); z.writestr("word/media/grey.png", png())


IN = 72.0
D9 = (9, 9, 9, 9)
LEFT = ("margin", 0)
TOP = ("paragraph", 0)

# Square: where the anchor run is, what is before the anchor paragraph.
sq = lambda: picture(2 * IN, IN, LEFT, TOP, dist=D9)
for where in ("start", "mid", "end"):
    write(f"sq-{where}", A + b_para(sq(), where) + C)
write("sq-end-heading", A + HEADING + b_para(sq(), "end") + C)
write("sq-start-heading", A + HEADING + b_para(sq(), "start") + C)
write("sq-end-keepnext", para(run(words("a", 40, 2)), "<w:keepNext/>") + b_para(sq(), "end") + C)

# Square: which lines the box takes, its bottom and top edges, with and without distances.
for f in (-20, -5, 5, 20):
    tag = ("m" if f < 0 else "p") + f"{abs(f):02d}"
    write(f"sq-h4{tag}", A + b_para(picture(2 * IN, 4 * LINE + f / 10, LEFT, TOP)) + C)
    write(f"sq-top-{tag}", A + b_para(picture(2 * IN, IN, LEFT, ("paragraph", 3 * LINE + f / 10))) + C)
write("sq-distb18", A + b_para(picture(2 * IN, 4 * LINE, LEFT, TOP, dist=(0, 18, 0, 0))) + C)
write("sq-distt18", A + b_para(picture(2 * IN, IN, LEFT, ("paragraph", 4 * LINE), dist=(18, 0, 0, 0))) + C)

# Square: the sides text goes on.
for side in ("bothSides", "left", "right", "largest"):
    short = {"bothSides": "both"}.get(side, side)
    write(f"side-{short}-c", A + b_para(picture(1.5 * IN, IN, ("margin", "center"), TOP, side=side, dist=D9)) + C)
for side in ("bothSides", "largest"):
    short = {"bothSides": "both"}.get(side, side)
    write(f"side-{short}-l", A + b_para(picture(1.5 * IN, IN, ("margin", IN), TOP, side=side, dist=D9)) + C)
    write(f"side-{short}-r", A + b_para(picture(1.5 * IN, IN, ("margin", 3.5 * IN), TOP, side=side, dist=D9)) + C)

# Square: the least room beside it that text goes in.
for g in (18, 36, 54, 72, 90, 108):
    write(f"gap-{g}", A + b_para(picture(2 * IN, IN, ("margin", g), TOP)) + C)
    write(f"gapr-{g}", A + b_para(picture(2 * IN, IN, ("margin", 6.5 * IN - g - 2 * IN), TOP)) + C)

# Top and bottom.
tb = lambda pv=TOP, dist=(0, 0, 0, 0): picture(2 * IN, IN, LEFT, pv, wrap="topAndBottom", dist=dist)
write("tb-start", A + b_para(tb(), "start") + C)
write("tb-end", A + b_para(tb(), "end") + C)
for k in (5, 25, 30):
    write(f"tb-off{k:02d}", A + b_para(tb(("paragraph", k / 10 * LINE))) + C)
write("tb-dist", A + b_para(tb(("paragraph", 2.5 * LINE), (9, 18, 0, 0))) + C)
write("tb-shape", A + b_para(shape(5.4 * IN, IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN), wrap="topAndBottom"), "end") + C)
# B's first line 3 lines above the page's bottom: 648pt of text, 48 lines above it.
write("tb-low", lines_of(44) + b_para(picture(2 * IN, 2 * IN, LEFT, TOP, wrap="topAndBottom"), "start") + C)

# Columns: B at the top of the second column.
COLBREAK = '<w:p><w:pPr>' + SINGLE + '</w:pPr><w:r><w:br w:type="column"/></w:r></w:p>'
for tag, ph in (("colright", ("column", "right")), ("coloff", ("column", 36)), ("marginoff", ("margin", 36)),
                ("marginright", ("margin", "right")), ("pageoff", ("page", 324))):
    write(f"col-{tag}", A + COLBREAK + b_para(picture(1.5 * IN, IN, ph, TOP, dist=D9), count=120) + C, columns=2)
write("col-cross", para(picture(1.5 * IN, IN, ("margin", 300), TOP, dist=D9) + run(words("a", 40, 2))) + b_para("", count=300) + C, columns=2)
# The first column: 40 lines of A, then B with a 2.6cm shape 0.2cm down it, as the magazine's pull quote.
write("col-tb-low", lines_of(37) + b_para(shape(216, 2.6 / 2.54 * IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN),
                                                                 wrap="topAndBottom"), "end", count=120) + C, columns=2)

# Batch 2. The same in Word 2010's layout (compatibilityMode 14, as python-docx's template has it - the
# showcase documents): does a line above the object move down to it?
C14 = {"sq-start": A + b_para(sq(), "start") + C,
       "sq-end-heading": A + HEADING + b_para(sq(), "end") + C,
       "side-both-c": A + b_para(picture(1.5 * IN, IN, ("margin", "center"), TOP, dist=D9)) + C,
       "tb-off25": A + b_para(tb(("paragraph", 2.5 * LINE))) + C,
       "gap-36": A + b_para(picture(2 * IN, IN, ("margin", 36), TOP)) + C}
for name, body in C14.items():
    write(f"{name}-c14", body, compat=14)
write("col-tb-low-c14", lines_of(37) + b_para(shape(216, 2.6 / 2.54 * IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN),
                                                                     wrap="topAndBottom"), "end", count=120) + C, columns=2, compat=14)
# The least room: words "il" (5pt) in a gap of g pt.
for g in (4, 8, 12, 16):
    write(f"gapil-{g}", A + para(picture(2 * IN, IN, ("margin", g), TOP) + run(" ".join(["il"] * 700))) + C)
# A paragraph across a page break with its picture's anchor at its start or end (B's first line 5 lines
# above the bottom of the page: A is 49 lines), and the shape of col-tb-low anchored at B's start.
LONG_A = lines_of(44)
for where in ():
    write(f"split-sq-{where}", LONG_A + b_para(picture(2 * IN, IN, LEFT, TOP), where) + C)
    write(f"split-tb-{where}", LONG_A + b_para(picture(2 * IN, 40, LEFT, TOP, wrap="topAndBottom"), where) + C)
write("col-tb-low-start", lines_of(37) + b_para(shape(216, 2.6 / 2.54 * IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN),
                                                                      wrap="topAndBottom"), "start", count=120) + C, columns=2)

# Batch 3. The least room, finer: "il" words in gaps of 18 ... 33pt.
for g in (18, 21, 24, 27, 30, 33):
    write(f"gapil-{g}", A + para(picture(2 * IN, IN, ("margin", g), TOP) + run(" ".join(["il"] * 700))) + C)
# Across a page break: A 44 lines, B's first line 4 lines above the page's bottom.
for where in ("start", "end"):
    write(f"split2-sq-{where}", LONG_A + b_para(picture(2 * IN, IN, LEFT, TOP), where) + C)
    write(f"split2-tb-{where}", LONG_A + b_para(picture(2 * IN, 40, LEFT, TOP, wrap="topAndBottom"), where) + C)
# The showcase's diagrams.docx "On site": Segoe UI 10pt at 1.15 lines, 6pt after; five empty paragraphs;
# a Georgia 14pt heading (10pt before, 4pt after, keepNext, keepLines); the picture (170 x 113.4pt, 9pt
# round) at the end of the justified paragraph after it. Word moved the heading down beside the picture.
def rep_styles(line=276, keep_lines=True):
    sp = f'<w:spacing w:after="120" w:line="{line}" w:lineRule="auto"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="{line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr>{sp}</w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:cs="Segoe UI"/><w:sz w:val="20"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
            f'<w:pPr><w:keepNext/>{"<w:keepLines/>" if keep_lines else ""}<w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
            '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>')
def rep_body(empties=5):
    plain = lambda content, after=None: f'<w:p><w:pPr>{"" if after is None else f"<w:spacing w:after={chr(34)}{after}{chr(34)}/>"}</w:pPr>{content}</w:p>'
    text = ("The crane that lifts the nacelle is the largest piece of equipment on site, and it is booked months ahead. "
            "If the wind is above 12 metres per second at hub height the lift waits; a week of bad weather can push the "
            "whole schedule into the next month, because the crane has another site to go to. Project managers keep a "
            "spreadsheet of forecast wind speeds for every lift window, and the most experienced of them say the "
            "spreadsheet is the most important document on the project.")
    body = plain(run("A callout marks the one stage that can stop everything after it."), 80)
    body += "".join(plain("", 0) for _ in range(empties + 1))
    body += '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>' + run("On site") + '</w:p>'
    body += '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>' + run(text) + picture(170.08, 113.39, LEFT, TOP, dist=D9) + '</w:p>'
    body += plain(run("Behind the text: a shape placed behind it, the text running over it, as a watermark does."), 80)
    return body
write("rep", rep_body(), compat=14, styles=rep_styles())
write("rep-c15", rep_body(), compat=15, styles=rep_styles())
write("rep-nokeeplines", rep_body(), compat=14, styles=rep_styles(keep_lines=False))
write("rep-single", rep_body(), compat=14, styles=rep_styles(line=240))
write("rep-noempty", rep_body(0), compat=14, styles=rep_styles())

# Batch 4. The least room to 0.5pt, and on the right; a keepLines paragraph above the picture in Word
# 2010's layout (rep: the heading moved down to the picture's top).
for g in (16.5, 17, 17.5):
    write(f"gapil-{g * 10:.0f}t", A + para(picture(2 * IN, IN, ("margin", g), TOP) + run(" ".join(["il"] * 700))) + C)
for g in (16, 18):
    write(f"gapilr-{g}", A + para(picture(2 * IN, IN, ("margin", 6.5 * IN - g - 2 * IN), TOP) + run(" ".join(["il"] * 700))) + C)
KEPT1 = para(run("Kept"), "<w:keepLines/>")
KEPT2 = para(run(words("k", 30, 2)), "<w:keepLines/>")
for compat in (14, 15):
    write(f"kl-one-c{compat}", A + KEPT1 + b_para(sq(), "end") + C, compat=compat)
    write(f"kl-two-c{compat}", A + KEPT2 + b_para(sq(), "end") + C, compat=compat)
    write(f"kl-none-c{compat}", A + para(run("Kept")) + b_para(sq(), "end") + C, compat=compat)
    write(f"kl-tb-c{compat}", A + KEPT1 + b_para(tb(dist=(9, 0, 0, 0)), "end") + C, compat=compat)

# Batch 5. The least room in older layouts: compatibilityMode 14 and 12, and a document without one
# (as tests/repros/word/float-image-wrap-square.docx: no text in its 34pt band).
for compat in (14, 12, None):
    for g in (16, 18, 24, 36, 48, 72):
        tag = f"c{compat}" if compat else "nocompat"
        write(f"gapil-{g}-{tag}", A + para(picture(2 * IN, IN, ("margin", g), TOP) + run(" ".join(["il"] * 700))) + C, compat=compat)
# Batch 6. The magazine's pull quote: a shape (216 x 73.7pt, 0.2cm down its paragraph, top and bottom)
# anchored at the end (or start) of a 6-line paragraph B that starts k lines down the first of two
# columns (k one-line paragraphs above it; a column holds 48): does B stay, go below the shape, go on in
# the second column whole or in part, and where is the shape? In Word 2010's and 2019's layouts.
B6 = lambda where: b_para(shape(216, 2.6 / 2.54 * IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN), wrap="topAndBottom"), where, count=48)
for compat in (14, 15):
    for k in (30, 36, 37, 39, 41, 42, 43):
        write(f"colq-k{k}-end-c{compat}", lines_of(k) + B6("end") + C, columns=2, compat=compat)
    for k in (39, 42, 43):
        write(f"colq-k{k}-start-c{compat}", lines_of(k) + B6("start") + C, columns=2, compat=compat)
# A column break's paragraph mark: text ending in a column break (B after it, nothing before / 12pt
# before), a paragraph that is only a column break with 12pt after, with its mark 20pt (w:pPr/w:rPr), with
# its break's run 20pt, and two such paragraphs in a row.
def colbreak(text="", ppr="", mark="", brk=""):
    return f'<w:p><w:pPr>{ppr}{SINGLE}{mark}</w:pPr>{run(text) if text else ""}<w:r>{brk}<w:br w:type="column"/></w:r></w:p>'
B_PROBE = para(run(words("b", 40, 3)))
B_PROBE12 = ('<w:p><w:pPr><w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
             + run(words("b", 40, 3)) + '</w:p>')
for compat in (14, 15):
    write(f"cbr-text-c{compat}", A + colbreak(words("t", 5, 2)) + B_PROBE, columns=2, compat=compat)
    write(f"cbr-text-b12-c{compat}", A + colbreak(words("t", 5, 2)) + B_PROBE12, columns=2, compat=compat)
    write(f"cbr-own-a12-c{compat}", A + colbreak().replace(SINGLE, '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"/>') + B_PROBE, columns=2, compat=compat)
    write(f"cbr-own-mark20-c{compat}", A + colbreak(mark='<w:rPr><w:sz w:val="40"/></w:rPr>') + B_PROBE, columns=2, compat=compat)
    write(f"cbr-own-run20-c{compat}", A + colbreak(brk='<w:rPr><w:sz w:val="40"/></w:rPr>') + B_PROBE, columns=2, compat=compat)
    write(f"cbr-two-c{compat}", A + colbreak() + colbreak() + B_PROBE, columns=2, compat=compat)
# Batch 7. Two lines' room below the shape (k40); and the showcase's magazine-feature.docx itself (its pull
# quote's paragraph would start 2 lines above the first column's foot below the quote: Word 2010's layout
# moved it to the second column whole) in Word 2019's layout, the quote at the paragraph's top (3 lines'
# room), 60pt / 47pt high (3 / 4 lines' room).
for compat in (14, 15):
    write(f"colq-k40-end-c{compat}", lines_of(40) + B6("end") + C, columns=2, compat=compat)
def magazine(name, compat=14, offset=None, height=None, part_edits=()):
    source = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../../../tests/test-documents-showcase/magazine-feature.docx")
    with zipfile.ZipFile(source) as z:
        parts = {item: z.read(item) for item in z.namelist()}
    doc = parts["word/document.xml"].decode("utf-8")
    if offset is not None:
        doc = doc.replace('<wp:positionV relativeFrom="paragraph"><wp:posOffset>72000</wp:posOffset>',
                          f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{round(offset * EMU)}</wp:posOffset>', 1)  # (the quote's: the first)
    if height is not None:
        doc = doc.replace('cx="2736000" cy="936000"', f'cx="2736000" cy="{round(height * EMU)}"')
    parts["word/document.xml"] = doc.encode("utf-8")
    for part, old, new in part_edits:
        parts[part] = parts[part].decode("utf-8").replace(old, new, 1).encode("utf-8")
    settings = parts["word/settings.xml"].decode("utf-8")
    parts["word/settings.xml"] = settings.replace('w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"',
                                                  f'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"').encode("utf-8")
    with zipfile.ZipFile(os.path.join(out, name + ".docx"), "w", zipfile.ZIP_DEFLATED) as z:
        for item, data in parts.items():
            z.writestr(item, data)
for compat in (14, 15):
    magazine(f"mag-c{compat}", compat)
    magazine(f"mag-off0-c{compat}", compat, offset=0)
    magazine(f"mag-h60-c{compat}", compat, height=60)
    magazine(f"mag-h47-c{compat}", compat, height=47)
# The second column full of text where the shape would be across in it (C of 800 words): does Word 2010's
# layout then move B whole, as in the magazine?
C_LONG = para(run(words("c", 800, 3)))
for compat in (14, 15):
    for k in (39, 40):
        write(f"colq-k{k}-full-c{compat}", lines_of(k) + B6("end") + C_LONG, columns=2, compat=compat)
# What in the magazine makes Word 2010's layout move the paragraph whole: without automatic hyphenation,
# with single line spacing (Normal's 1.15 gone), without the paragraph's first-line indent.
magazine("mag-nohyph-c14", part_edits=[("word/settings.xml", "<w:autoHyphenation/>", "")])
magazine("mag-single-c14", part_edits=[("word/styles.xml", 'w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia"', 'w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia"')])
magazine("mag-noind-c14", part_edits=[("word/document.xml", '<w:ind w:firstLine="283"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t>The engineers', '<w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t>The engineers')])
# ... with the quote as wide as the column (224.1pt), and with the paragraph's own space after 0.
magazine("mag-wide-c14", part_edits=[("word/document.xml", 'cx="2736000" cy="936000"', 'cx="2846070" cy="936000"'),
                                     ("word/document.xml", 'cx="2736000" cy="936000"', 'cx="2846070" cy="936000"')])
magazine("mag-after0-c14", part_edits=[("word/document.xml", '<w:spacing w:after="120"/><w:ind w:firstLine="283"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t>The engineers', '<w:spacing w:after="0"/><w:ind w:firstLine="283"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t>The engineers')])
# Its space after 0 the magazine's paragraph goes on in the first column (2 lines) in Word 2010's layout too:
# B with space after (6pt: its 2 lines and that past the column's foot; 4pt: within it), and without the shape.
def b_after(obj, after, where="end"):
    return b_para(obj, where, count=48).replace(SINGLE, f'<w:spacing w:before="0" w:after="{after * 20}" w:line="240" w:lineRule="auto"/>', 1)
SHAPE6 = lambda: shape(216, 2.6 / 2.54 * IN, ("column", 0), ("paragraph", 0.2 / 2.54 * IN), wrap="topAndBottom")
for compat in (14, 15):
    write(f"colq-k40-a6-c{compat}", lines_of(40) + b_after(SHAPE6(), 6) + C, columns=2, compat=compat)
    write(f"colq-k40-a4-c{compat}", lines_of(40) + b_after(SHAPE6(), 4) + C, columns=2, compat=compat)
    write(f"colq-k46-a6-noshape-c{compat}", lines_of(46) + b_after("", 6) + C, columns=2, compat=compat)
print(out)
