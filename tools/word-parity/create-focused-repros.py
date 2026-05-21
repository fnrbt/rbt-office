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


def paragraph(
    text="",
    style=None,
    alignment=None,
    spacing_before=None,
    spacing_after=None,
    list_id=None,
    list_level=None,
    indent_left=None,
    hanging=None,
    first_line=None,
    bold=False,
    paragraph_bold=False,
    preserve_space=True,
):
    props = []
    if style:
        props.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if alignment:
        props.append(f'<w:jc w:val="{esc(alignment)}"/>')
    if paragraph_bold:
        props.append("<w:rPr><w:b/></w:rPr>")
    spacing_attrs = []
    if spacing_before is not None:
        spacing_attrs.append(f'w:before="{spacing_before}"')
    if spacing_after is not None:
        spacing_attrs.append(f'w:after="{spacing_after}"')
    if spacing_attrs:
        props.append(f"<w:spacing {' '.join(spacing_attrs)}/>")
    if list_id is not None:
        level = 0 if list_level is None else list_level
        props.append(
            "<w:numPr>"
            f'<w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{list_id}"/>'
            "</w:numPr>"
        )
    indent_attrs = []
    if indent_left is not None:
        indent_attrs.append(f'w:left="{indent_left}"')
    if hanging is not None:
        indent_attrs.append(f'w:hanging="{hanging}"')
    if first_line is not None:
        indent_attrs.append(f'w:firstLine="{first_line}"')
    if indent_attrs:
        props.append(f"<w:ind {' '.join(indent_attrs)}/>")
    ppr = f"<w:pPr>{''.join(props)}</w:pPr>" if props else ""
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    space = ' xml:space="preserve"' if preserve_space else ""
    return (
        f"<w:p>{ppr}<w:r>{rpr}<w:t{space}>{esc(text)}</w:t></w:r></w:p>"
    )


def empty_paragraph():
    return "<w:p/>"


def empty_paragraph_with_run_props():
    return '<w:p><w:pPr><w:rPr><w:lang w:val="en-US"/></w:rPr></w:pPr></w:p>'


def page_break_paragraph():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def cell(
    text,
    width=2400,
    grid_span=None,
    vertical_merge=None,
    spacing_after=None,
    padding_left=None,
    padding_top=None,
    padding_right=None,
    padding_bottom=None,
    alignment=None,
    valign=None,
    shading=None,
    shading_color="auto",
    shading_fill="auto",
    bold=False,
    paragraph_bold=False,
    preserve_space=True,
):
    tc_props = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if grid_span is not None:
        tc_props.append(f'<w:gridSpan w:val="{grid_span}"/>')
    if vertical_merge is not None:
        if vertical_merge == "continue":
            tc_props.append("<w:vMerge/>")
        else:
            tc_props.append(f'<w:vMerge w:val="{esc(vertical_merge)}"/>')
    if shading:
        tc_props.append(
            f'<w:shd w:val="{esc(shading)}" w:color="{esc(shading_color)}" w:fill="{esc(shading_fill)}"/>'
        )
    if valign:
        tc_props.append(f'<w:vAlign w:val="{esc(valign)}"/>')
    padding = []
    if padding_top is not None:
        padding.append(f'<w:top w:w="{padding_top}" w:type="dxa"/>')
    if padding_left is not None:
        padding.append(f'<w:left w:w="{padding_left}" w:type="dxa"/>')
    if padding_bottom is not None:
        padding.append(f'<w:bottom w:w="{padding_bottom}" w:type="dxa"/>')
    if padding_right is not None:
        padding.append(f'<w:right w:w="{padding_right}" w:type="dxa"/>')
    if padding:
        tc_props.append(f"<w:tcMar>{''.join(padding)}</w:tcMar>")
    return (
        "<w:tc>"
        f"<w:tcPr>{''.join(tc_props)}</w:tcPr>"
        f"{paragraph(text, alignment=alignment, spacing_after=spacing_after, bold=bold, paragraph_bold=paragraph_bold, preserve_space=preserve_space)}"
        "</w:tc>"
    )


def cell_with_table(intro, nested_table_xml, width=4606, trailing=""):
    """A table cell whose content is an intro paragraph, a nested table, then a
    closing paragraph (Word requires a paragraph after a nested table)."""
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f"{paragraph(intro)}"
        f"{nested_table_xml}"
        f"{paragraph(trailing)}"
        "</w:tc>"
    )


def table_border(side, val, sz=None, color="auto"):
    attrs = [f'w:val="{esc(val)}"']
    if sz is not None:
        attrs.append(f'w:sz="{sz}"')
    if color is not None:
        attrs.append(f'w:color="{esc(color)}"')
    return f"<w:{side} {' '.join(attrs)}/>"


def cell_borders(**sides):
    ordered = ["top", "left", "bottom", "right", "insideH", "insideV"]
    parts = []
    for side in ordered:
        value = sides.get(side)
        if value is None:
            continue
        if isinstance(value, str):
            parts.append(table_border(side, value, None, None))
        else:
            parts.append(table_border(side, *value))
    return f"<w:tcBorders>{''.join(parts)}</w:tcBorders>" if parts else ""


def empty_cell_with_props(width=2303, grid_span=None, vertical_merge=None, borders_xml=""):
    tc_props = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if grid_span is not None:
        tc_props.append(f'<w:gridSpan w:val="{grid_span}"/>')
    if vertical_merge is not None:
        if vertical_merge == "continue":
            tc_props.append("<w:vMerge/>")
        else:
            tc_props.append(f'<w:vMerge w:val="{esc(vertical_merge)}"/>')
    if borders_xml:
        tc_props.append(borders_xml)
    return (
        "<w:tc>"
        f"<w:tcPr>{''.join(tc_props)}</w:tcPr>"
        "<w:p/>"
        "</w:tc>"
    )


def cell_paragraphs(paragraphs, width=2400, padding_left=None):
    tc_props = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if padding_left is not None:
        tc_props.append(f'<w:tcMar><w:left w:w="{padding_left}" w:type="dxa"/></w:tcMar>')
    return (
        "<w:tc>"
        f"<w:tcPr>{''.join(tc_props)}</w:tcPr>"
        f"{''.join(paragraphs)}"
        "</w:tc>"
    )


def row(values, spacing_after=None, height=None):
    tr_props = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height is not None else ""
    return "<w:tr>" + tr_props + "".join(cell(value, spacing_after=spacing_after) for value in values) + "</w:tr>"


def row_cells(cells):
    return "<w:tr>" + "".join(cells) + "</w:tr>"


def row_cells_props(cells, height=None):
    tr_props = f'<w:trPr><w:trHeight w:val="{height}"/></w:trPr>' if height is not None else ""
    return "<w:tr>" + tr_props + "".join(cells) + "</w:tr>"


def header_row_cells(cells):
    return "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + "".join(cells) + "</w:tr>"


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


def table_with_grid(rows, grid_widths, table_width="0", width_type="auto"):
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblW w:w="{table_width}" w:type="{width_type}"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def table_grid_style_table(rows, grid_widths, table_width="0", width_type="auto"):
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="{table_width}" w:type="{width_type}"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def table_grid_style_table_with_pr(rows, grid_widths, extra_tbl_pr="", table_width="0", width_type="auto"):
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="{table_width}" w:type="{width_type}"/>'
        f"{extra_tbl_pr}"
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def named_style_table(style_id, rows, grid_widths, table_width="0", width_type="auto"):
    grid = "".join(f'<w:gridCol w:w="{grid_width}"/>' for grid_width in grid_widths)
    return (
        "<w:tbl>"
        f'<w:tblPr><w:tblStyle w:val="{esc(style_id)}"/><w:tblW w:w="{table_width}" w:type="{width_type}"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        "</w:tblPr>"
        f"<w:tblGrid>{grid}</w:tblGrid>"
        + "".join(row_values for row_values in rows)
        + "</w:tbl>"
    )


def exact_height_spacer_table(height, width=9212):
    return (
        "<w:tbl>"
        '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>'
        '<w:tr><w:trPr>'
        f'<w:trHeight w:val="{height}" w:hRule="exact"/>'
        '</w:trPr>'
        '<w:tc><w:tcPr>'
        f'<w:tcW w:w="{width}" w:type="dxa"/>'
        '</w:tcPr><w:p/></w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )


def border_stress_table():
    return table_grid_style_table([
        row_cells([
            empty_cell_with_props(2303, borders_xml=cell_borders(right=("single", 8, "17220B"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(top=("single", 8, "17220B"), left=("single", 8, "17220B"), bottom=("single", 8, "17220B"), right=("single", 8, "17220B"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(top=("single", 24, "17220B"), left=("single", 8, "17220B"), bottom=("single", 24, "17220B"), right=("single", 24, "17220B"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(top=("single", 48, "17220B"), left=("single", 24, "17220B"), bottom=("single", 48, "17220B"), right=("single", 48, "17220B"))),
        ]),
        row_cells([
            empty_cell_with_props(2303),
            empty_cell_with_props(4606, grid_span=2, vertical_merge="restart"),
            empty_cell_with_props(2303, borders_xml=cell_borders(top=("single", 48, "17220B"), bottom=("double", 12, "auto"))),
        ]),
        row_cells([
            empty_cell_with_props(2303),
            empty_cell_with_props(4606, grid_span=2, vertical_merge="continue", borders_xml=cell_borders(right=("double", 12, "auto"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(top=("double", 12, "auto"), left=("double", 12, "auto"), bottom=("double", 12, "auto"), right=("double", 12, "auto"))),
        ]),
        row_cells([
            empty_cell_with_props(4606, grid_span=2, vertical_merge="restart", borders_xml=cell_borders(top=("dotted", 24, "auto"), left=("dotted", 24, "auto"), bottom=("dotted", 24, "auto"), right=("dotted", 24, "auto"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(left=("dotted", 24, "auto"))),
            empty_cell_with_props(2303, vertical_merge="restart", borders_xml=cell_borders(top=("double", 12, "auto"))),
        ]),
        row_cells([
            empty_cell_with_props(4606, grid_span=2, vertical_merge="continue", borders_xml=cell_borders(top="nil", left=("dotted", 24, "auto"), bottom=("dotted", 24, "auto"), right=("dotted", 24, "auto"))),
            empty_cell_with_props(2303, borders_xml=cell_borders(left=("dotted", 24, "auto"))),
            empty_cell_with_props(2303, vertical_merge="continue"),
        ]),
    ], [2303, 2303, 2303, 2303])


def repeated_long_paragraphs(count):
    text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
    )
    return "".join(paragraph(text) for _ in range(count))


def docx4j_page_geometry():
    return {
        "page_width": 11906,
        "page_height": 16838,
        "margin_top": 1417,
        "margin_right": 1417,
        "margin_bottom": 1134,
        "margin_left": 1417,
        "header": 708,
        "footer": 708,
        "doc_grid": '<w:docGrid w:linePitch="360"/>',
    }


def long_paragraph(spacing_before=None):
    text = (
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
    )
    return paragraph(text, spacing_before=spacing_before)


def section_properties(include_header_footer=False, margin_bottom=1440, include_doc_grid=False, page_geometry=None):
    refs = (
        '<w:headerReference w:type="default" r:id="rId3"/>'
        '<w:footerReference w:type="default" r:id="rId4"/>'
        if include_header_footer
        else ""
    )
    page_geometry = page_geometry or {}
    page_width = page_geometry.get("page_width", 12240)
    page_height = page_geometry.get("page_height", 15840)
    margin_top = page_geometry.get("margin_top", 1440)
    margin_right = page_geometry.get("margin_right", 1440)
    margin_bottom = page_geometry.get("margin_bottom", margin_bottom)
    margin_left = page_geometry.get("margin_left", 1440)
    header = page_geometry.get("header", 720)
    footer = page_geometry.get("footer", 720)
    doc_grid = page_geometry.get("doc_grid", '<w:docGrid w:type="default" w:linePitch="360" w:charSpace="4294965247"/>') if include_doc_grid else ""
    cols = page_geometry.get("cols", "")
    return f'<w:sectPr>{refs}<w:pgSz w:w="{page_width}" w:h="{page_height}"/><w:pgMar w:top="{margin_top}" w:right="{margin_right}" w:bottom="{margin_bottom}" w:left="{margin_left}" w:header="{header}" w:footer="{footer}"/>{cols}{doc_grid}</w:sectPr>'


def document(body_xml, leading_section_properties=False, include_header_footer=False, margin_bottom=1440, include_doc_grid=False, page_geometry=None):
    sectpr = section_properties(include_header_footer, margin_bottom, include_doc_grid, page_geometry)
    if leading_section_properties:
        body = f"{sectpr}{body_xml}"
    else:
        body = f"{body_xml}{sectpr}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}"><w:body>'
        f"{body}"
        "</w:body></w:document>"
    )


def styles(include_heading_spacing=True, include_doc_defaults=True):
    heading1_spacing = '<w:spacing w:after="240"/>' if include_heading_spacing else ""
    heading2_spacing = '<w:spacing w:before="240" w:after="160"/>' if include_heading_spacing else ""
    doc_defaults = (
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        if include_doc_defaults
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        f"{doc_defaults}"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        f'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr>{heading1_spacing}<w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:pPr>{heading2_spacing}<w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        "</w:styles>"
    )


def footnote_styles():
    # Defines FootnoteText (10pt, the Word default for footnote bodies) and the
    # superscript FootnoteReference char style, so Word renders the footnotes at
    # 10pt like a real document (rather than falling back to the 11pt body size).
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
        "</w:styles>"
    )


def table_grid_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="0" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def table_grid_paragraph_spacing_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def table_grid_dense_row_pitch_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def table_grid_line_spacing_only_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>{cell_margin}</w:tblPr></w:style>'
        f'<w:style w:type="table" w:styleId="LineOnlyTable"><w:name w:val="Line Only Table"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def docx4j_table_grid_styles():
    borders = "".join(
        f'<w:{name} w:val="single" w:sz="4" w:space="0" w:color="17220B" w:themeColor="background1" w:themeShade="1A"/>'
        for name in ["top", "left", "bottom", "right", "insideH", "insideV"]
    )
    cell_margin = (
        '<w:tblCellMar>'
        '<w:top w:w="0" w:type="dxa"/>'
        '<w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="{REL_NS}" xmlns:w="{WORD_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-DE" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
        f'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>{cell_margin}</w:tblPr></w:style>'
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
        f'<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>{cell_margin}</w:tblPr></w:style>'
        "</w:styles>"
    )


def top_heading_before_spacing_styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        "</w:style>"
        "</w:styles>"
    )


def cambria_heading_wrap_styles():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">'
        '<w:name w:val="Normal"/>'
        '<w:pPr><w:widowControl/><w:spacing w:lineRule="auto" w:line="276" w:before="100" w:after="100"/><w:ind w:left="567" w:hanging="0"/><w:jc w:val="left"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        "</w:style>"
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:ind w:left="567" w:hanging="567"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="365F91"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        "</w:style>"
        "</w:styles>"
    )


def content_types(include_numbering=False, include_header_footer=False, include_office_theme=False, include_footnotes=False, include_endnotes=False):
    footnotes_override = (
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        if include_footnotes
        else ""
    )
    endnotes_override = (
        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        if include_endnotes
        else ""
    )
    numbering_override = (
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        if include_numbering
        else ""
    )
    running_overrides = (
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        if include_header_footer
        else ""
    )
    theme_overrides = (
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
        '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        if include_office_theme
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{CONTENT_TYPES_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        f"{numbering_override}"
        f"{running_overrides}"
        f"{theme_overrides}"
        f"{footnotes_override}"
        f"{endnotes_override}"
        "</Types>"
    )


def root_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )


def document_rels(include_numbering=False, include_header_footer=False, include_office_theme=False, include_footnotes=False, include_endnotes=False):
    footnotes_rel = (
        '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
        if include_footnotes
        else ""
    )
    endnotes_rel = (
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>'
        if include_endnotes
        else ""
    )
    numbering_rel = (
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
        if include_numbering
        else ""
    )
    running_rels = (
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        if include_header_footer
        else ""
    )
    theme_rels = (
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
        '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        if include_office_theme
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        f"{numbering_rel}"
        f"{running_rels}"
        f"{theme_rels}"
        f"{footnotes_rel}"
        f"{endnotes_rel}"
        "</Relationships>"
    )


def footnote_paragraph(text_before, note_id, text_after=""):
    """A body paragraph that drops a footnote reference between two text runs."""
    ref_run = (
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>'
        f'<w:footnoteReference w:id="{note_id}"/></w:r>'
    )
    after_run = f'<w:r><w:t xml:space="preserve">{esc(text_after)}</w:t></w:r>' if text_after else ""
    return (
        "<w:p><w:r><w:t xml:space=\"preserve\">"
        f"{esc(text_before)}</w:t></w:r>{ref_run}{after_run}</w:p>"
    )


def footnotes_part(notes):
    """notes: list of (id, text). Includes the mandatory separator footnotes."""
    body = [
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for note_id, text in notes:
        body.append(
            f'<w:footnote w:id="{note_id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {esc(text)}</w:t></w:r></w:p></w:footnote>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{WORD_NS}">'
        + "".join(body)
        + "</w:footnotes>"
    )


def endnote_paragraph(text_before, note_id, text_after=""):
    """A body paragraph that drops an endnote reference between two text runs."""
    ref_run = (
        '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>'
        f'<w:endnoteReference w:id="{note_id}"/></w:r>'
    )
    after_run = f'<w:r><w:t xml:space="preserve">{esc(text_after)}</w:t></w:r>' if text_after else ""
    return (
        "<w:p><w:r><w:t xml:space=\"preserve\">"
        f"{esc(text_before)}</w:t></w:r>{ref_run}{after_run}</w:p>"
    )


def endnotes_part(notes):
    """notes: list of (id, text). Includes the mandatory separator endnotes."""
    body = [
        '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>',
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>',
    ]
    for note_id, text in notes:
        body.append(
            f'<w:endnote w:id="{note_id}"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {esc(text)}</w:t></w:r></w:p></w:endnote>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:endnotes xmlns:w="{WORD_NS}">'
        + "".join(body)
        + "</w:endnotes>"
    )


def endnote_styles():
    # Defines EndnoteText (10pt, Word's default for endnote bodies) and the
    # superscript EndnoteReference char style, mirroring footnote_styles.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        "<w:docDefaults>"
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        "</w:docDefaults>"
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
        "</w:styles>"
    )


def numbering():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering xmlns:w="{WORD_NS}">'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>'
        '<w:abstractNum w:abstractNumId="1">'
        '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/><w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="■"/><w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        "</w:numbering>"
    )


def header_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:hdr xmlns:w="{WORD_NS}">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Focused Header</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Page 1</w:t></w:r></w:p>'
        '</w:hdr>'
    )


def footer_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{WORD_NS}">'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>Focused Footer</w:t></w:r></w:p>'
        '</w:ftr>'
    )


def word_settings_xml(compatibility_mode=14):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{WORD_NS}">'
        '<w:zoom w:percent="100"/>'
        '<w:defaultTabStop w:val="720"/>'
        '<w:compat>'
        f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compatibility_mode}"/>'
        '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
        '</w:compat>'
        '<w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi=""/>'
        '</w:settings>'
    )


def docx4j_word_settings_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{WORD_NS}">'
        '<w:zoom w:percent="100"/>'
        '<w:defaultTabStop w:val="708"/>'
        '<w:hyphenationZone w:val="425"/>'
        '<w:characterSpacingControl w:val="doNotCompress"/>'
        '<w:compat>'
        '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>'
        '</w:compat>'
        '<w:themeFontLang w:val="de-DE" w:eastAsia="ko-KR"/>'
        '<w:decimalSymbol w:val="."/>'
        '<w:listSeparator w:val=","/>'
        '</w:settings>'
    )


def font_table_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts xmlns:w="{WORD_NS}" xmlns:r="{REL_NS}">'
        '<w:font w:name="Calibri"><w:charset w:val="01"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Cambria"><w:charset w:val="01"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Arial"><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font>'
        '<w:font w:name="Times New Roman"><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
        "</w:fonts>"
    )


def office_theme_xml(major_font="Cambria", minor_font="Calibri"):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
        '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
        '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
        '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
        '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
        '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
        '<a:accent6><a:srgbClr val="F79646"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Office">'
        f'<a:majorFont><a:latin typeface="{esc(major_font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        f'<a:minorFont><a:latin typeface="{esc(minor_font)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/>'
        '<a:extraClrSchemeLst/>'
        '</a:theme>'
    )


def theme_xml():
    return office_theme_xml()


def docx4j_theme_xml():
    return office_theme_xml("Arial", "Arial")


def write_docx(path, body_xml, include_numbering=False, leading_section_properties=False, include_header_footer=False, margin_bottom=1440, include_heading_spacing=True, include_doc_defaults=True, styles_xml=None, include_office_theme=False, include_doc_grid=False, page_geometry=None, compatibility_mode=14, settings_xml=None, theme_xml_value=None, footnotes_xml=None, endnotes_xml=None):
    include_footnotes = footnotes_xml is not None
    include_endnotes = endnotes_xml is not None
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "_rels").mkdir()
        (root / "word" / "_rels").mkdir(parents=True)
        (root / "[Content_Types].xml").write_text(content_types(include_numbering, include_header_footer, include_office_theme, include_footnotes, include_endnotes), encoding="utf-8")
        (root / "_rels" / ".rels").write_text(root_rels(), encoding="utf-8")
        (root / "word" / "document.xml").write_text(document(body_xml, leading_section_properties, include_header_footer, margin_bottom, include_doc_grid, page_geometry), encoding="utf-8")
        (root / "word" / "styles.xml").write_text(styles_xml or styles(include_heading_spacing, include_doc_defaults), encoding="utf-8")
        if include_numbering:
            (root / "word" / "numbering.xml").write_text(numbering(), encoding="utf-8")
        if include_header_footer:
            (root / "word" / "header1.xml").write_text(header_xml(), encoding="utf-8")
            (root / "word" / "footer1.xml").write_text(footer_xml(), encoding="utf-8")
        if include_office_theme:
            (root / "word" / "theme").mkdir()
            (root / "word" / "settings.xml").write_text(settings_xml or word_settings_xml(compatibility_mode), encoding="utf-8")
            (root / "word" / "fontTable.xml").write_text(font_table_xml(), encoding="utf-8")
            (root / "word" / "theme" / "theme1.xml").write_text(theme_xml_value or theme_xml(), encoding="utf-8")
        if include_footnotes:
            (root / "word" / "footnotes.xml").write_text(footnotes_xml, encoding="utf-8")
        if include_endnotes:
            (root / "word" / "endnotes.xml").write_text(endnotes_xml, encoding="utf-8")
        (root / "word" / "_rels" / "document.xml.rels").write_text(document_rels(include_numbering, include_header_footer, include_office_theme, include_footnotes, include_endnotes), encoding="utf-8")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in sorted(root.rglob("*")):
                if file.is_file():
                    info = zipfile.ZipInfo(file.relative_to(root).as_posix(), (1980, 1, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_DEFLATED
                    zf.writestr(info, file.read_bytes())


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("tests/repros/word")
    repros = {
        "table-import-preserves-structure.docx": (False, (
            paragraph("Table import repro", "Heading1")
            + table([row(["A1", "B1"]), row(["A2", "B2"])])
            + paragraph("After table")
        )),
        "table-cell-default-after-spacing.docx": (False, (
            paragraph("Table cell spacing repro", "Heading1")
            + table([row(["Alpha", "Beta"]), row(["Gamma", "Delta"])])
            + paragraph("After table")
        )),
        "table-cell-explicit-left-padding.docx": (False, (
            paragraph("Table cell explicit left padding repro", "Heading1")
            + table([
                row_cells([
                    cell("Name", padding_left=108),
                    cell("Value", padding_left=108),
                ]),
                row_cells([
                    cell("Country", padding_left=108),
                    cell("AU", padding_left=108),
                ]),
            ])
            + paragraph("After padded table")
        )),
        "table-dxa-grid-alignment-row-height.docx": (False, (
            paragraph("Table DXA grid alignment row height repro", "Heading1")
            + table_with_grid([
                row_cells_props([
                    cell("A", width=4606),
                    cell("B", width=4606, alignment="right"),
                ], height=1134),
                row_cells_props([
                    cell("C", width=4606),
                    cell("D", width=4606, alignment="right"),
                ], height=1134),
            ], [4606, 4606])
            + paragraph("After table")
        )),
        "table-cell-auto-pattern-shading.docx": (False, (
            paragraph("Table cell auto pattern shading repro", "Heading1")
            + table_with_grid([
                row_cells_props([
                    cell("pct5 auto", width=4606, shading="pct5"),
                    cell("pct10 auto", width=4606, shading="pct10"),
                ], height=720),
                row_cells_props([
                    cell("plain", width=4606),
                    cell("green stripe", width=4606, shading="reverseDiagStripe", shading_color="6FB31A"),
                ], height=720),
            ], [4606, 4606])
            + paragraph("After shaded table")
        )),
        "table-grid-style-borders.docx": (False, False, False, (
            paragraph("TableGrid style border repro", "Heading1")
            + table_grid_style_table([
                row_cells([
                    cell("A", width=3000),
                    cell("B", width=3000),
                ]),
                row_cells([
                    cell("C", width=3000),
                    cell("D", width=3000),
                ]),
            ], [3000, 3000])
            + paragraph("After styled table")
        ), True, True, table_grid_styles()),
        "table-style-paragraph-spacing.docx": (False, False, False, (
            paragraph("Table style paragraph spacing repro", "Heading1")
            + table_grid_style_table([
                row_cells([cell("row 01 a", width=2303), cell("row 01 b", width=2303), cell("row 01 c", width=2303), cell("row 01 d", width=2303)]),
                row_cells([cell("row 02 a", width=2303), cell("row 02 b", width=2303), cell("row 02 c", width=2303), cell("row 02 d", width=2303)]),
                row_cells([cell("row 03 a", width=2303), cell("row 03 b", width=2303), cell("row 03 c", width=2303), cell("row 03 d", width=2303)]),
                row_cells([cell("row 04 a", width=2303), cell("row 04 b", width=2303), cell("row 04 c", width=2303), cell("row 04 d", width=2303)]),
            ], [2303, 2303, 2303, 2303])
            + paragraph("After table")
        ), True, True, table_grid_paragraph_spacing_styles()),
        "table-style-cell-margin-origin.docx": (False, False, False, (
            paragraph("Table style cell margin origin repro", "Heading1")
            + table_grid_style_table([
                row_cells([cell("Alpha", width=2303), cell("Beta", width=2303)]),
                row_cells([cell("Gamma", width=2303), cell("Delta", width=2303)]),
            ], [2303, 2303])
            + paragraph("After styled margin table")
        ), True, True, table_grid_paragraph_spacing_styles()),
        "table-grid-dense-row-pitch.docx": (False, False, False, (
            paragraph("TableGrid dense row pitch repro", "Heading1")
            + table_grid_style_table([
                row_cells([
                    cell(f"row {index:02d} a", width=2303),
                    cell(f"row {index:02d} b", width=2303),
                    cell(f"row {index:02d} c", width=2303),
                    cell(f"row {index:02d} d", width=2303),
                ])
                for index in range(1, 15)
            ], [2303, 2303, 2303, 2303])
            + paragraph("After dense table")
        ), True, True, table_grid_dense_row_pitch_styles()),
        "table-style-line-spacing-only.docx": (False, False, False, (
            paragraph("Table style line-spacing-only repro", "Heading1")
            + named_style_table("LineOnlyTable", [
                row_cells([cell("Alpha", width=2303), cell("Beta", width=2303)]),
                row_cells([cell("Gamma", width=2303), cell("Delta", width=2303)]),
                row_cells([cell("Epsilon", width=2303), cell("Zeta", width=2303)]),
            ], [2303, 2303])
            + paragraph("After dense line-only table")
        ), True, True, table_grid_line_spacing_only_styles()),
        "table-header-dense-row-pitch.docx": (False, False, False, (
            paragraph("Headers")
            + table_grid_style_table([
                header_row_cells([
                    cell("Header1ab", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center"),
                    cell("Header1cd", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center"),
                ]),
                header_row_cells([
                    cell("Header2a", width=2303, shading="pct5", shading_fill="auto", alignment="center"),
                    cell("Header2b", width=2303, shading="pct5", shading_fill="auto", alignment="center"),
                    cell("Header2c", width=2303, shading="pct5", shading_fill="auto", alignment="center"),
                    cell("Header2d", width=2303, shading="pct5", shading_fill="auto", alignment="center"),
                ]),
                *[
                    row_cells([
                        cell(f"row {index:02d} a", width=2303),
                        cell(f"row {index:02d} b", width=2303),
                        cell(f"row {index:02d} c", width=2303),
                        cell(f"row {index:02d} d", width=2303),
                    ])
                    for index in range(1, 15)
                ],
            ], [2303, 2303, 2303, 2303])
            + paragraph("After header dense table")
        ), True, True, table_grid_dense_row_pitch_styles(), True, True, docx4j_page_geometry(), 12, docx4j_word_settings_xml()),
        "table-header-docx4j-style-row-pitch.docx": (False, False, False, (
            paragraph("Headers", preserve_space=False)
            + table_grid_style_table([
                header_row_cells([
                    cell("Header1ab", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center", bold=True, paragraph_bold=True, preserve_space=False),
                    cell("Header1cd", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center", bold=True, paragraph_bold=True, preserve_space=False),
                ]),
                header_row_cells([
                    cell("Header2a", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2b", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2c", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2d", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                ]),
                *[
                    row_cells([
                        cell(f"row {index:02d} a", width=2303, preserve_space=False),
                        cell(f"row {index:02d} b", width=2303, preserve_space=False),
                        cell(f"row {index:02d} c", width=2303, preserve_space=False),
                        cell(f"row {index:02d} d", width=2303, preserve_space=False),
                    ])
                    for index in range(1, 15)
                ],
            ], [2303, 2303, 2303, 2303])
            + paragraph("After header dense table", preserve_space=False)
        ), True, True, docx4j_table_grid_styles(), True, True, docx4j_page_geometry(), 12, docx4j_word_settings_xml(), docx4j_theme_xml()),
        "table-header-page2-docx4j-row-pitch.docx": (False, False, False, (
            page_break_paragraph()
            + exact_height_spacer_table(1020)
            + paragraph("Headers", preserve_space=False)
            + table_grid_style_table([
                header_row_cells([
                    cell("Header1ab", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center", bold=True, paragraph_bold=True, preserve_space=False),
                    cell("Header1cd", width=4606, grid_span=2, shading="pct5", shading_fill="auto", alignment="center", bold=True, paragraph_bold=True, preserve_space=False),
                ]),
                header_row_cells([
                    cell("Header2a", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2b", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2c", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                    cell("Header2d", width=2303, shading="pct5", shading_fill="auto", alignment="center", preserve_space=False),
                ]),
                *[
                    row_cells([
                        cell(f"row {index:02d} a", width=2303, preserve_space=False),
                        cell(f"row {index:02d} b", width=2303, preserve_space=False),
                        cell(f"row {index:02d} c", width=2303, preserve_space=False),
                        cell(f"row {index:02d} d", width=2303, preserve_space=False),
                    ])
                    for index in range(1, 15)
                ],
            ], [2303, 2303, 2303, 2303])
            + paragraph("After header dense table", preserve_space=False)
        ), True, True, docx4j_table_grid_styles(), True, True, docx4j_page_geometry(), 12, docx4j_word_settings_xml(), docx4j_theme_xml()),
        "table-gridspan-bottom-fit.docx": (False, False, False, (
            paragraph("Table gridspan bottom-fit repro", "Heading1")
            + page_break_paragraph()
            + exact_height_spacer_table(12900)
            + paragraph("Mergin, Auto size")
            + table_grid_style_table([
                row_cells([
                    cell("A", width=6141, grid_span=2),
                    cell("B", width=3071),
                ]),
                row_cells([
                    cell("C", width=3070),
                    cell("D", width=6142, grid_span=2),
                ]),
            ], [3070, 3071, 3071])
            + paragraph("After gridspan table")
        ), True, True, table_grid_dense_row_pitch_styles(), False, False, docx4j_page_geometry()),
        "table-spaced-sequence-page-break.docx": (False, False, False, (
            paragraph("Table spaced sequence page-break repro", preserve_space=False)
            + page_break_paragraph()
            + exact_height_spacer_table(9650)
            + paragraph("Shading, Margins, Cell spacing", preserve_space=False)
            + table_grid_style_table_with_pr([
                row_cells([
                    cell("Cell 1a", width=2303, shading="clear", shading_fill="auto", preserve_space=False),
                    cell("Cell 1b", width=2303, shading="clear", shading_fill="auto", preserve_space=False),
                    cell("Cell 1c", width=2303, shading="clear", shading_fill="auto", preserve_space=False),
                    cell("Cell 1d", width=2303, shading="clear", shading_fill="auto", preserve_space=False),
                ]),
                row_cells([
                    cell("Cell 2a", width=2303, shading="pct5", shading_fill="FFC000", preserve_space=False),
                    cell("Cell 2b", width=2303, shading="pct5", shading_fill="FFC000", preserve_space=False),
                    cell("Cell 2c", width=2303, shading="pct5", shading_fill="FFC000", preserve_space=False),
                    cell("Cell 2d", width=2303, shading="pct5", shading_fill="FFC000", padding_top=284, padding_right=567, padding_bottom=284, padding_left=567, preserve_space=False),
                ]),
            ], [2583, 2440, 2440, 2651],
                '<w:tblCellSpacing w:w="142" w:type="dxa"/>'
                '<w:shd w:val="clear" w:color="auto" w:fill="FFC000"/>'
                '<w:tblCellMar><w:top w:w="113" w:type="dxa"/><w:left w:w="227" w:type="dxa"/><w:bottom w:w="113" w:type="dxa"/><w:right w:w="227" w:type="dxa"/></w:tblCellMar>')
            + empty_paragraph()
            + paragraph("Mergin, Auto size", preserve_space=False)
            + table_grid_style_table([
                row_cells([
                    cell("A", width=6141, grid_span=2, preserve_space=False),
                    cell("B", width=3071, preserve_space=False),
                ]),
                row_cells([
                    cell("C", width=3070, preserve_space=False),
                    cell("D", width=6142, grid_span=2, preserve_space=False),
                ]),
            ], [3070, 3071, 3071])
            + empty_paragraph_with_run_props()
            + paragraph("Auto size", preserve_space=False)
            + table_grid_style_table([
                row_cells([
                    cell("A", width=375, preserve_space=False),
                    cell("B", width=375, preserve_space=False),
                ]),
                row_cells([
                    cell("C", width=375, preserve_space=False),
                    cell("D", width=375, preserve_space=False),
                ]),
            ], [375, 375])
        ), True, True, docx4j_table_grid_styles(), True, True, docx4j_page_geometry(), 12, docx4j_word_settings_xml(), docx4j_theme_xml()),
        "table-border-empty-cell-height.docx": (False, False, False, (
            paragraph("Border empty-cell height repro", preserve_space=False)
            + paragraph("Borders", preserve_space=False)
            + border_stress_table()
            + empty_paragraph_with_run_props()
            + paragraph("Indent (WW uses gridBefore/gridAfter as above)", preserve_space=False)
            + table_grid_style_table([
                row_cells([
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1536, grid_span=2),
                    empty_cell_with_props(1536, grid_span=2),
                ]),
                row_cells([
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1535, grid_span=2),
                    empty_cell_with_props(1536, grid_span=2),
                    empty_cell_with_props(1536, grid_span=2),
                ]),
            ], [708, 827, 708, 827, 708, 827, 708, 827, 708, 828, 708, 828, 708], table_width="9920", width_type="dxa")
        ), True, True, docx4j_table_grid_styles(), True, True, docx4j_page_geometry(), 12, docx4j_word_settings_xml(), docx4j_theme_xml()),
        "table-valign-row-height.docx": (False, False, False, (
            paragraph("Table valign row height repro", "Heading1")
            + table_grid_style_table([
                row_cells_props([
                    cell("A", width=4606),
                    cell("B", width=4606, alignment="right"),
                ], height=1134),
                row_cells_props([
                    cell("C", width=4606, valign="bottom"),
                    cell("D", width=4606, alignment="right", valign="bottom"),
                ], height=1134),
            ], [4606, 4606])
            + paragraph("After valign table")
        ), True, True, table_grid_paragraph_spacing_styles()),
        "table-vmerge-valign-rowspan.docx": (False, False, False, (
            paragraph("Table vMerge valign row-span repro", "Heading1")
            + table_grid_style_table([
                row_cells_props([
                    cell("A", width=4606),
                    cell("B", width=4606, vertical_merge="restart", valign="center"),
                ], height=1134),
                row_cells_props([
                    cell("C", width=4606),
                    empty_cell_with_props(4606, vertical_merge="continue"),
                ], height=1134),
            ], [4606, 4606])
            + paragraph("After vMerge table")
        ), True, True, table_grid_paragraph_spacing_styles()),
        "table-cell-spacing-margins.docx": (False, False, False, (
            paragraph("Table cell spacing and margins repro", spacing_after=0)
            + table_grid_style_table_with_pr([
                row_cells([
                    cell("Cell 1a", width=2303, shading="clear", shading_fill="auto"),
                    cell("Cell 1b", width=2303, shading="clear", shading_fill="auto"),
                    cell("Cell 1c", width=2303, shading="clear", shading_fill="auto"),
                    cell("Cell 1d", width=2303, shading="clear", shading_fill="auto"),
                ]),
                row_cells([
                    cell("Cell 2a", width=2303, shading="pct5", shading_fill="FFC000"),
                    cell("Cell 2b", width=2303, shading="pct5", shading_fill="FFC000"),
                    cell("Cell 2c", width=2303, shading="pct5", shading_fill="FFC000"),
                    cell("Cell 2d", width=2303, shading="pct5", shading_fill="FFC000", padding_top=284, padding_right=567, padding_bottom=284, padding_left=567),
                ]),
            ], [2583, 2440, 2440, 2651],
                '<w:tblCellSpacing w:w="142" w:type="dxa"/>'
                '<w:shd w:val="clear" w:color="auto" w:fill="FFC000"/>'
                '<w:tblCellMar><w:top w:w="113" w:type="dxa"/><w:left w:w="227" w:type="dxa"/><w:bottom w:w="113" w:type="dxa"/><w:right w:w="227" w:type="dxa"/></w:tblCellMar>')
            + paragraph("After spaced margin table")
        ), True, True, table_grid_paragraph_spacing_styles(), False, False, docx4j_page_geometry()),
        "table-cell-explicit-before-spacing.docx": (False, (
            paragraph("Table cell before-spacing repro", "Heading1")
            + table([
                row_cells([
                    cell_paragraphs([paragraph("First", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([paragraph("", spacing_before=100, spacing_after=0)]),
                ]),
                row_cells([
                    cell_paragraphs([paragraph("Second", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                        paragraph("", spacing_before=100, spacing_after=0),
                    ]),
                ]),
                row_cells([
                    cell_paragraphs([paragraph("Third", spacing_before=100, spacing_after=0)]),
                    cell_paragraphs([paragraph("", spacing_before=100, spacing_after=0)]),
                ]),
            ])
            + paragraph("After table")
        )),
        "nested-table-cell-layout.docx": (False, (
            paragraph("Nested table cell layout repro", "Heading1")
            + table_with_grid([
                row_cells([
                    cell("Outer left cell", width=4606),
                    cell_with_table(
                        "Outer right intro",
                        table_with_grid([
                            row_cells([cell("N1", width=2000), cell("N2", width=2000)]),
                            row_cells([cell("N3", width=2000), cell("N4", width=2000)]),
                        ], [2000, 2000]),
                        width=4606,
                        trailing="Outer right after nested",
                    ),
                ]),
                row_cells([
                    cell("Outer bottom left", width=4606),
                    cell("Outer bottom right", width=4606),
                ]),
            ], [4606, 4606])
            + paragraph("After outer table")
        )),
        "two-tables-border-position.docx": (False, (
            paragraph("Two table border position repro", "Heading1")
            + paragraph("First table", "Heading2")
            + table([row(["A1", "B1", "C1"]), row(["A2", "B2", "C2"])])
            + paragraph("\n")
            + paragraph("Second table", "Heading2")
            + table([row(["Product", "Price", "Quantity", "Total"]), row(["Widget A", "$10.00", "5", "$50.00"]), row(["Total", "", "", "$50.00"])])
            + paragraph("After second table")
        )),
        "normal-paragraph-line-pitch.docx": (False, (
            paragraph("Normal paragraph pitch repro", "Heading1")
            + paragraph("First normal line")
            + paragraph("Second normal line")
            + paragraph("Third normal line")
            + paragraph("Wrapped line pitch - Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        )),
        "alignment-inherits-default-spacing.docx": (False, (
            paragraph("Alignment inherits default spacing repro", "Heading1")
            + paragraph("Before aligned paragraph")
            + paragraph("Left aligned direct pPr", alignment="left")
            + paragraph("Center aligned direct pPr", alignment="center")
            + paragraph("Right aligned direct pPr", alignment="right")
            + paragraph("After aligned paragraph")
        )),
        "empty-normal-paragraph-mark-size.docx": (False, False, False, (
            paragraph("Empty paragraph mark size repro")
            + empty_paragraph()
            + paragraph("Second line after an empty Normal paragraph")
            + empty_paragraph()
            + paragraph("Third line after another empty Normal paragraph")
        ), True, False),
        "heading-style-no-spacing.docx": (False, False, False, (
            paragraph("Heading style no spacing repro", "Heading1")
            + paragraph("")
            + paragraph("Section A", "Heading2")
            + paragraph("First normal line")
            + paragraph("Second normal line")
            + paragraph("")
            + paragraph("Section B", "Heading2")
            + paragraph("Third normal line")
        ), False),
        "top-heading-before-spacing.docx": (False, False, False, (
            paragraph("Top heading before-spacing repro", "Heading1")
            + paragraph("Second heading with the same imported before-spacing", "Heading1")
            + paragraph("After heading")
        ), True, True, top_heading_before_spacing_styles()),
        "heading-cambria-font-wrap.docx": (True, False, False, (
            paragraph("Heading Cambria font wrap repro", spacing_before=0, spacing_after=0, indent_left=0, hanging=0)
            + paragraph("Subject to the terms and conditions of this Agreement, You hereby grant to the Sponsor and to recipients of software distributed by the Sponsor a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable (except as stated in this section) patent license to make, have made, use, offer to sell, sell, import, and otherwise transfer the Work, where such license applies only to those patent claims licensable by You that are necessarily infringed by Your Contribution(s) alone or by combination of Your Contribution(s) with the Work to which such Contribution(s) was submitted. If any entity institutes patent litigation against You or any other entity (including a cross-claim or counterclaim in a lawsuit) alleging that your Contribution, or the Work to which you have contributed, constitutes direct or contributory patent infringement, then any patent licenses granted to that entity under this Agreement for that Contribution or Work shall terminate as of the date such litigation is filed. ")
            + paragraph("You represent that you are legally entitled to grant the above license. ", "Heading1", list_id=1, list_level=0, indent_left=567, hanging=567)
            + paragraph("If your employer(s) has rights to intellectual property that you create that includes your Contributions, you represent that you have received permission to make Contributions on behalf of that employer, that your employer has waived such rights for your Contributions to the Sponsor, or that your employer has executed a separate Corporate CLA with the Sponsor. ")
        ), True, True, cambria_heading_wrap_styles(), True, True),
        "wrapped-paragraph-line-pitch.docx": (False, (
            paragraph("Wrapped paragraph line pitch repro", "Heading1")
            + repeated_long_paragraphs(18)
            + paragraph("After wrapped paragraphs")
        )),
        "widow-control-last-line.docx": (False, (
            paragraph("Widow control last line repro", "Heading1")
            + repeated_long_paragraphs(12)
            + long_paragraph()
            + paragraph("After widow probe")
        ), 1740),
        "list-default-indent.docx": (True, False, (
            paragraph("List default indent repro", "Heading1")
            + paragraph("Top bullet", list_id=2, list_level=0)
            + paragraph("Nested bullet", list_id=2, list_level=1)
            + paragraph("Deep bullet", list_id=2, list_level=2)
            + paragraph("Top number", list_id=1, list_level=0)
            + paragraph("Nested number", list_id=1, list_level=1)
            + paragraph("Deep number", list_id=1, list_level=2)
            + paragraph("After list")
        )),
        "running-header-body-offset.docx": (False, True, True, (
            paragraph("Running header body offset repro", "Heading1")
            + paragraph("First body paragraph")
            + paragraph("Second body paragraph")
        )),
    }
    for name, value in repros.items():
        margin_bottom = 1440
        include_heading_spacing = True
        include_doc_defaults = True
        include_office_theme = False
        include_doc_grid = False
        page_geometry = None
        theme_xml_value = None
        if len(value) == 2:
            include_numbering, body = value
            leading_section_properties = False
            include_header_footer = False
            styles_xml = None
            compatibility_mode = 14
            settings_xml = None
            theme_xml_value = None
        elif len(value) == 3:
            styles_xml = None
            compatibility_mode = 14
            settings_xml = None
            theme_xml_value = None
            if isinstance(value[2], int):
                include_numbering, body, margin_bottom = value
                leading_section_properties = False
                include_header_footer = False
            else:
                include_numbering, leading_section_properties, body = value
                include_header_footer = False
        else:
            include_numbering, leading_section_properties, include_header_footer, body = value[:4]
            styles_xml = None
            compatibility_mode = 14
            settings_xml = None
            theme_xml_value = None
            if len(value) > 4:
                include_heading_spacing = value[4]
            if len(value) > 5:
                include_doc_defaults = value[5]
            if len(value) > 6:
                styles_xml = value[6]
            if len(value) > 7:
                include_office_theme = value[7]
            if len(value) > 8:
                include_doc_grid = value[8]
            if len(value) > 9:
                page_geometry = value[9]
            if len(value) > 10:
                compatibility_mode = value[10]
            if len(value) > 11:
                settings_xml = value[11]
            if len(value) > 12:
                theme_xml_value = value[12]
        path = out_dir / name
        write_docx(path, body, include_numbering, leading_section_properties, include_header_footer, margin_bottom, include_heading_spacing, include_doc_defaults, styles_xml, include_office_theme, include_doc_grid, page_geometry, compatibility_mode, settings_xml, theme_xml_value)
        print(path)

    # Footnote placement repro (needs the footnotes.xml part, so it bypasses the
    # tuple-based table above and calls write_docx directly).
    footnote_body = (
        paragraph("Footnote placement repro", "Heading1")
        + footnote_paragraph(
            "This sentence ends with a footnote reference",
            1,
            " and then continues on so the line wraps normally afterward.")
        + repeated_long_paragraphs(6)
        + footnote_paragraph("A second reference appears here", 2, " near the bottom of the body text.")
        + paragraph("Final body paragraph after the second reference.")
    )
    footnote_path = out_dir / "footnote-placement.docx"
    write_docx(
        footnote_path,
        footnote_body,
        styles_xml=footnote_styles(),
        footnotes_xml=footnotes_part([
            (1, "First footnote: Word renders this in a smaller font at the bottom of the page above the body bottom margin."),
            (2, "Second footnote: it sits just below the first, separated from the body by the footnote separator line."),
        ]),
    )
    print(footnote_path)

    # Endnote repro (needs the endnotes.xml part). Endnotes collect at the end of
    # the document, after all body content, with a separator above them.
    endnote_body = (
        paragraph("Endnote placement repro", "Heading1")
        + endnote_paragraph(
            "This sentence ends with an endnote reference",
            1,
            " and then continues on so the line wraps normally afterward.")
        + repeated_long_paragraphs(4)
        + endnote_paragraph("A second reference appears here", 2, " in the body text.")
        + paragraph("Final body paragraph after the second reference.")
    )
    endnote_path = out_dir / "endnote-placement.docx"
    write_docx(
        endnote_path,
        endnote_body,
        styles_xml=endnote_styles(),
        endnotes_xml=endnotes_part([
            (1, "First endnote: Word collects endnotes at the end of the document in a smaller font."),
            (2, "Second endnote: it sits just below the first, after the endnote separator line."),
        ]),
    )
    print(endnote_path)

    # Two-column section: body text flows down column 1 then column 2 (Word
    # balances the last page; here there is enough text to fill both columns).
    twocol_body = (
        paragraph("Two column section repro", "Heading1")
        + repeated_long_paragraphs(10)
    )
    twocol_path = out_dir / "two-column-section.docx"
    write_docx(
        twocol_path,
        twocol_body,
        page_geometry={"cols": '<w:cols w:num="2" w:space="720"/>'},
    )
    print(twocol_path)

    twocol_sep_path = out_dir / "two-column-separator.docx"
    write_docx(
        twocol_sep_path,
        paragraph("Two column with separator", "Heading1") + repeated_long_paragraphs(10),
        page_geometry={"cols": '<w:cols w:num="2" w:space="720" w:sep="1"/>'},
    )
    print(twocol_sep_path)


if __name__ == "__main__":
    main()
