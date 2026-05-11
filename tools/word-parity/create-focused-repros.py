#!/usr/bin/env python3
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def esc(text):
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def paragraph(text="", style=None, spacing_after=None):
    props = []
    if style:
        props.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if spacing_after is not None:
        props.append(f'<w:spacing w:after="{spacing_after}"/>')
    ppr = f"<w:pPr>{''.join(props)}</w:pPr>" if props else ""
    return (
        f"<w:p>{ppr}<w:r><w:t xml:space=\"preserve\">{esc(text)}</w:t></w:r></w:p>"
    )


def cell(text, width=2400, spacing_after=None):
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f"{paragraph(text, spacing_after=spacing_after)}"
        "</w:tc>"
    )


def row(values, spacing_after=None):
    return "<w:tr>" + "".join(cell(value, spacing_after=spacing_after) for value in values) + "</w:tr>"


def table(rows, width_pct=5000):
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblW w:w="{width_pct}" w:type="pct"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def document(body_xml):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}"><w:body>'
        f"{body_xml}"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
        "</w:body></w:document>"
    )


def styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="160"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        "</w:styles>"
    )


def content_types():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CONTENT_TYPES_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        "</Types>"
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )


def document_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )


def write_docx(path, body_xml):
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "_rels").mkdir()
        (root / "word" / "_rels").mkdir(parents=True)
        (root / "[Content_Types].xml").write_text(content_types(), encoding="utf-8")
        (root / "_rels" / ".rels").write_text(root_rels(), encoding="utf-8")
        (root / "word" / "document.xml").write_text(document(body_xml), encoding="utf-8")
        (root / "word" / "styles.xml").write_text(styles(), encoding="utf-8")
        (root / "word" / "_rels" / "document.xml.rels").write_text(document_rels(), encoding="utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(root.rglob("*")):
                if file.is_file():
                    zf.write(file, file.relative_to(root).as_posix())


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("tests/repros/word")
    repros = {
        "table-import-preserves-structure.docx": (
            paragraph("Table import repro", "Heading1")
            + table([row(["A1", "B1"]), row(["A2", "B2"])])
            + paragraph("After table")
        ),
        "table-cell-default-after-spacing.docx": (
            paragraph("Table cell spacing repro", "Heading1")
            + table([row(["Alpha", "Beta"]), row(["Gamma", "Delta"])])
            + paragraph("After table")
        ),
        "two-tables-border-position.docx": (
            paragraph("Two table border position repro", "Heading1")
            + paragraph("First table", "Heading2")
            + table([row(["A1", "B1", "C1"]), row(["A2", "B2", "C2"])])
            + paragraph("\n")
            + paragraph("Second table", "Heading2")
            + table([row(["Product", "Price", "Quantity", "Total"]), row(["Widget A", "$10.00", "5", "$50.00"]), row(["Total", "", "", "$50.00"])])
            + paragraph("After second table")
        ),
        "normal-paragraph-line-pitch.docx": (
            paragraph("Normal paragraph pitch repro", "Heading1")
            + paragraph("First normal line")
            + paragraph("Second normal line")
            + paragraph("Third normal line")
            + paragraph("Wrapped line pitch - Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        ),
    }
    for name, body in repros.items():
        path = out_dir / name
        write_docx(path, body)
        print(path)


if __name__ == "__main__":
    main()
