#!/usr/bin/env python3
"""Construct a .docx of DrawingML autoshapes (a:prstGeom / a:custGeom / a:gradFill)
to validate the preset-geometry shape renderer. -> tests/repros/word-shapes/"""
import os, zipfile

OUT = os.path.join(os.path.dirname(__file__), "../../tests/repros/word-shapes")
os.makedirs(OUT, exist_ok=True)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = (f'xmlns:w="{W}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:wps="{WPS}" xmlns:r="{R}"')
EMU = 9525  # per px

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>'

def fill_solid(hexv):
    return f'<a:solidFill><a:srgbClr val="{hexv}"/></a:solidFill>'

def fill_grad(c0, c1, ang_deg):
    ang = int(ang_deg * 60000)
    return (f'<a:gradFill><a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="{c0}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{c1}"/></a:gs>'
            f'</a:gsLst><a:lin ang="{ang}" scaled="1"/></a:gradFill>')

def spPr(px, py, w, h, geom, fill, line="2F5496"):
    return (f'<wps:spPr><a:xfrm><a:off x="{px*EMU}" y="{py*EMU}"/><a:ext cx="{w*EMU}" cy="{h*EMU}"/></a:xfrm>'
            f'{geom}{fill}<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln></wps:spPr>')

def custom_spPr(px, py, w, h, fill):
    # a chevron-ish arrow via custGeom in a 100x100 coord space
    path = ('<a:custGeom><a:pathLst><a:path w="100" h="100">'
            '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
            '<a:lnTo><a:pt x="70" y="0"/></a:lnTo>'
            '<a:lnTo><a:pt x="100" y="50"/></a:lnTo>'
            '<a:lnTo><a:pt x="70" y="100"/></a:lnTo>'
            '<a:lnTo><a:pt x="0" y="100"/></a:lnTo>'
            '<a:lnTo><a:pt x="30" y="50"/></a:lnTo>'
            '<a:close/></a:path></a:pathLst></a:custGeom>')
    return (f'<wps:spPr><a:xfrm><a:off x="{px*EMU}" y="{py*EMU}"/><a:ext cx="{w*EMU}" cy="{h*EMU}"/></a:xfrm>'
            f'{path}{fill}<a:ln w="12700"><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln></wps:spPr>')

shape_n = 0
def anchor(px, py, w, h, sppr, txt=None):
    global shape_n
    shape_n += 1
    txbx = ''
    if txt:
        txbx = (f'<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">{txt}</w:t></w:r></w:p></w:txbxContent></wps:txbx>')
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{shape_n}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{px*EMU}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{py*EMU}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{w*EMU}" cy="{h*EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{shape_n}" name="Shape {shape_n}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{WPS}"><wps:wsp><wps:cNvSpPr/>{sppr}{txbx}<wps:bodyPr/>'
            f'</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')

# grid of shapes: 3 columns
cols = [60, 230, 400]
W_, H_ = 140, 90
rows = [
    ("ellipse", fill_solid("4472C4")),
    ("roundRect", fill_solid("70AD47")),
    ("triangle", fill_solid("ED7D31")),
    ("diamond", fill_solid("FFC000")),
    ("star5", fill_solid("FFD966")),
    ("rightArrow", fill_solid("5B9BD5")),
    ("pentagon", fill_solid("A5A5A5")),
    ("hexagon", fill_solid("9E480E")),
    ("chevron", fill_solid("C55A11")),
]
anchors = []
for i, (geom, fill) in enumerate(rows):
    col = cols[i % 3]
    row_y = 60 + (i // 3) * 130
    g = f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
    anchors.append(anchor(col, row_y, W_, H_, spPr(col, row_y, W_, H_, g, fill), txt=geom))

# gradient ellipse + custom-path chevron on a 4th row
y4 = 60 + 3 * 130
anchors.append(anchor(cols[0], y4, W_, H_, spPr(cols[0], y4, W_, H_,
              '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>', fill_grad("FFE599", "BF9000", 0)), txt="gradient"))
anchors.append(anchor(cols[1], y4, W_, H_, custom_spPr(cols[1], y4, W_, H_, fill_solid("8FAADC")), txt="custom"))

body = f'<w:p><w:pPr></w:pPr><w:r><w:t>Shape geometry gallery</w:t></w:r>{"".join(anchors)}</w:p>'
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{body}{SECT}</w:body></w:document>'

path = os.path.join(OUT, "autoshapes-gallery.docx")
with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("word/document.xml", doc)
print("wrote", path, "-", shape_n, "shapes")
