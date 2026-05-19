#!/usr/bin/env python3
"""Generate focused DOCX repros for floating-image wrap parity.

Separate from create-focused-repros.py because the image-bearing packages
need extra parts (media files, image relationships, png content-type) that
the simple text-only writer there does not handle.
"""

import struct
import sys
import tempfile
import zipfile
import zlib
from pathlib import Path

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_INCH = 914400
EMU_PER_PIXEL = 9525  # at 96dpi, matches reader's emuToPixels


def make_solid_png(width: int, height: int, rgb=(204, 64, 64)) -> bytes:
    """Build a minimal RGB PNG of given size filled with one color."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8-bit RGB
    row = b"\x00" + bytes(rgb) * width  # filter byte 0 + RGB pixels
    raw = row * height
    idat = zlib.compress(raw, 9)
    return sig + chunk(b"IHDR", ihdr_data) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


def anchored_image_xml(rid: str, name: str, *, width_emu: int, height_emu: int,
                       pos_h_emu: int, pos_v_emu: int,
                       wrap: str = "wrapSquare",
                       relative_from_h: str = "page",
                       relative_from_v: str = "page",
                       behind_doc: str = "0",
                       allow_overlap: str = "1") -> str:
    """One run carrying a w:drawing with a wp:anchor (floating image)."""
    if wrap == "wrapSquare":
        wrap_xml = '<wp:wrapSquare wrapText="bothSides"/>'
    elif wrap == "wrapTopAndBottom":
        wrap_xml = '<wp:wrapTopAndBottom/>'
    elif wrap == "wrapNone":
        wrap_xml = '<wp:wrapNone/>'
    else:
        raise ValueError(f"unsupported wrap: {wrap}")
    return (
        '<w:r>'
        '<w:drawing>'
        f'<wp:anchor xmlns:wp="{WP_NS}" '
        'distT="0" distB="0" distL="114300" distR="114300" '
        'simplePos="0" relativeHeight="251658240" '
        f'behindDoc="{behind_doc}" locked="0" layoutInCell="1" '
        f'allowOverlap="{allow_overlap}">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{relative_from_h}"><wp:posOffset>{pos_h_emu}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{relative_from_v}"><wp:posOffset>{pos_v_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{width_emu}" cy="{height_emu}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'{wrap_xml}'
        f'<wp:docPr id="1" name="{name}"/>'
        '<wp:cNvGraphicFramePr/>'
        f'<a:graphic xmlns:a="{A_NS}">'
        f'<a:graphicData uri="{PIC_NS}">'
        f'<pic:pic xmlns:pic="{PIC_NS}">'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="1" name="{name}"/>'
        '<pic:cNvPicPr/>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip xmlns:r="{R_NS}" r:embed="{rid}"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
        '</w:r>'
    )


def text_run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


WRAP_PARAGRAPH_TEXT = (
    "This paragraph should wrap around the anchored image on its right side. "
    "Word treats the floating picture as an obstacle and shortens the line "
    "extent on the lines whose vertical span overlaps the image rectangle. "
    "We need enough body text here to force several wrapped lines so we can "
    "measure where the wrap edge sits, how far the obstacle pushes the line "
    "start or end, and whether the line height inside the wrap band differs "
    "from the unobstructed line height below the image. Add a second long "
    "sentence with very ordinary words to keep the line breaks predictable "
    "across renderers, since unusual punctuation tends to introduce small "
    "kerning differences that drown out the actual layout deltas we care "
    "about. After the image clears, lines should resume the full body width."
)


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}">'
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>'
        '</w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr>'
        '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
        '</w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/>'
        '<w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:sz w:val="32"/><w:color w:val="2F5496"/></w:rPr>'
        '</w:style>'
        '</w:styles>'
    )


def document_xml(body_paragraphs_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" '
        f'xmlns:r="{R_NS}" '
        f'xmlns:wp="{WP_NS}" '
        f'xmlns:a="{A_NS}" '
        f'xmlns:pic="{PIC_NS}">'
        '<w:body>'
        f'{body_paragraphs_xml}'
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
        '</w:body>'
        '</w:document>'
    )


def content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CONTENT_TYPES_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )


def document_rels_xml(image_rels) -> str:
    image_xml = "".join(
        f'<Relationship Id="{rid}" Type="{R_NS}/image" Target="{target}"/>'
        for rid, target in image_rels
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        f'{image_xml}'
        '</Relationships>'
    )


def heading(text: str) -> str:
    return (
        '<w:p>'
        '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        '</w:p>'
    )


def write_docx(path: Path, body_paragraphs_xml: str, images) -> None:
    """Write a DOCX with optional embedded PNGs.

    images: list of (rid, target_path_in_zip, png_bytes).
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "_rels").mkdir()
        (root / "word" / "_rels").mkdir(parents=True)
        (root / "word" / "media").mkdir(parents=True, exist_ok=True)

        (root / "[Content_Types].xml").write_text(content_types_xml(), encoding="utf-8")
        (root / "_rels" / ".rels").write_text(root_rels_xml(), encoding="utf-8")
        (root / "word" / "document.xml").write_text(document_xml(body_paragraphs_xml), encoding="utf-8")
        (root / "word" / "styles.xml").write_text(styles_xml(), encoding="utf-8")

        image_rels = []
        for rid, target, png_bytes in images:
            image_rels.append((rid, target))
            full = root / "word" / target
            full.parent.mkdir(parents=True, exist_ok=True)
            full.write_bytes(png_bytes)

        (root / "word" / "_rels" / "document.xml.rels").write_text(
            document_rels_xml(image_rels), encoding="utf-8"
        )

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(root.rglob("*")):
                if file.is_file():
                    info = zipfile.ZipInfo(file.relative_to(root).as_posix(),
                                           (1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    zf.writestr(info, file.read_bytes())


def build_wrap_square_repro() -> tuple[str, str, list]:
    """First float repro: a wrapSquare image anchored to a body paragraph.

    Geometry:
      - 8.5"x11" letter, 1" margins, so body width is 6.5" = 5,943,600 EMU.
      - Image is 1.6"x1.2" (1,463,040 x 1,097,280 EMU) anchored 1.6" right of
        page-left and 1.7" below page-top, putting it inside the body box
        with text wrapping on both sides.
    """
    image_width_emu = int(1.6 * EMU_PER_INCH)
    image_height_emu = int(1.2 * EMU_PER_INCH)
    pos_h_emu = int(1.6 * EMU_PER_INCH)
    pos_v_emu = int(1.7 * EMU_PER_INCH)

    drawing_run = anchored_image_xml(
        "rId100", "Float repro image",
        width_emu=image_width_emu,
        height_emu=image_height_emu,
        pos_h_emu=pos_h_emu,
        pos_v_emu=pos_v_emu,
        wrap="wrapSquare",
    )

    body = (
        heading("Float wrap-square repro")
        + '<w:p>'
        + drawing_run
        + text_run(WRAP_PARAGRAPH_TEXT)
        + '</w:p>'
        + '<w:p>'
        + text_run(
            "A short trailing paragraph confirms that the body resumes the "
            "full available width once the image has cleared above."
        )
        + '</w:p>'
    )

    png = make_solid_png(width=240, height=180, rgb=(204, 64, 64))
    images = [("rId100", "media/image1.png", png)]
    return "float-image-wrap-square.docx", body, images


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("tests/repros/word")
    repros = [build_wrap_square_repro()]
    for name, body, images in repros:
        path = out_dir / name
        write_docx(path, body, images)
        print(path)


if __name__ == "__main__":
    main()
