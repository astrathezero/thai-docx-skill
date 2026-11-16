#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
thai_docx.py - Production-ready Thai Word Document (.docx) Generator & Formatter

Solves OpenXML Thai typography challenges in Microsoft Word (Windows & macOS):
1. Thai Distributed justification (w:val="thaiDistribute") prevents huge gaps
   between words and prevents awkward character-level stretching.
2. Full Complex Script (CS) tagging: w:cs, w:szCs, w:bCs, w:iCs, w:hint="cs".
3. Full Thai language tagging: w:lang val="th-TH", eastAsia="th-TH", bidi="th-TH".
4. Thai Government & Standard Typography (TH Sarabun New 16pt, hanging indents,
   numerals normalization, no colons in section headers).
5. High-quality table borders, zebra striping, and cell margins.
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DEFAULT_FONT = "TH Sarabun New"
DEFAULT_LANG = "th-TH"
DEFAULT_BODY_SIZE_PT = 16
DEFAULT_LINE_SPACING = 1.15

THAI_DIGITS = "๐๑๒๓๔๕๖๗๘๙"
ARABIC_DIGITS = "0123456789"
THAI_TO_ARABIC_TRANS = str.maketrans(THAI_DIGITS, ARABIC_DIGITS)
ARABIC_TO_THAI_TRANS = str.maketrans(ARABIC_DIGITS, THAI_DIGITS)


# =============================================================================
# 1. TEXT CLEANING & NUMERAL UTILITIES
# =============================================================================

def thai_to_arabic(text: str) -> str:
    """Convert Thai numerals (๐-๙) to Arabic numerals (0-9)."""
    if not text:
        return ""
    return text.translate(THAI_TO_ARABIC_TRANS)


def arabic_to_thai(text: str) -> str:
    """Convert Arabic numerals (0-9) to Thai numerals (๐-๙)."""
    if not text:
        return ""
    return text.translate(ARABIC_TO_THAI_TRANS)


def remove_colons(text: str) -> str:
    """Remove English colons (:) often erroneously used after Thai section names or words like ได้แก่:"""
    if not text:
        return ""
    return text.replace(":", "")


def clean_thai_text(text: str, to_arabic: bool = True, remove_colon: bool = True) -> str:
    """Normalize Thai text according to document preferences."""
    if not text:
        return ""
    res = text
    if to_arabic:
        res = thai_to_arabic(res)
    if remove_colon:
        res = remove_colons(res)
    return res


# =============================================================================
# 2. OPENXML COMPLEX SCRIPT & LANGUAGE TAGGING
# =============================================================================

def set_run_font(
    run,
    name: str = DEFAULT_FONT,
    size_pt: float = DEFAULT_BODY_SIZE_PT,
    bold: bool = False,
    italic: bool = False,
    color_rgb: RGBColor = None,
    lang_code: str = DEFAULT_LANG
):
    """
    Applies complete Complex Script (CS) font and language properties to a Run.
    Essential for Microsoft Word (macOS & Windows) to render Thai without bugs.
    """
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb

    rPr = run._element.get_or_add_rPr()

    # 1. Set rFonts across all font families + hint="cs"
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ["ascii", "hAnsi", "cs", "eastAsia"]:
        rFonts.set(qn(f"w:{attr}"), name)
    rFonts.set(qn("w:hint"), "cs")

    # 2. Set font sizes in half-points (both Western w:sz and Complex Script w:szCs)
    sz_val = str(int(size_pt * 2))
    for tag_name in ["w:sz", "w:szCs"]:
        elem = rPr.find(qn(tag_name))
        if elem is None:
            elem = OxmlElement(tag_name)
            rPr.append(elem)
        elem.set(qn("w:val"), sz_val)

    # 3. Complex Script flag <w:cs/>
    if rPr.find(qn("w:cs")) is None:
        rPr.append(OxmlElement("w:cs"))

    # 4. Complex Script bold <w:bCs/> and italic <w:iCs/>
    if bold:
        if rPr.find(qn("w:bCs")) is None:
            rPr.append(OxmlElement("w:bCs"))
    if italic:
        if rPr.find(qn("w:iCs")) is None:
            rPr.append(OxmlElement("w:iCs"))

    # 5. Language tagging for proofing and line-breaking
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), lang_code)


def set_paragraph_thai(
    p,
    alignment: str = "thaiDistribute",
    space_before_pt: float = 0,
    space_after_pt: float = 4,
    line_spacing: float = DEFAULT_LINE_SPACING,
    first_line_indent_in: float = 0,
    left_indent_in: float = 0,
    keep_with_next: bool = False,
    lang_code: str = DEFAULT_LANG
):
    """
    Configures paragraph layout with proper Thai alignment and OpenXML language tagging.
    
    Alignment options:
      - 'thaiDistribute': (Default) Thai Distributed (กระจายแบบไทย). Evens out spacing across
        Thai word clusters without giant gaps.
      - 'both' / 'justify': Standard Western justification (adjusts spaces between words only).
      - 'left', 'center', 'right'.
    """
    pPr = p._p.get_or_add_pPr()

    # Alignment
    if alignment in ["thaiDistribute", "thai_justify", "thai"]:
        p.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    elif alignment in ["both", "justify"]:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif alignment == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alignment == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif alignment == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Spacing and Indents
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.keep_with_next = keep_with_next

    if first_line_indent_in != 0:
        p.paragraph_format.first_line_indent = Inches(first_line_indent_in)
    if left_indent_in != 0:
        p.paragraph_format.left_indent = Inches(left_indent_in)

    # Paragraph-level rPr (sets default run properties for paragraph marks)
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)

    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), lang_code)
    lang.set(qn("w:eastAsia"), lang_code)
    lang.set(qn("w:bidi"), lang_code)

    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:hint"), "cs")

    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), "32")  # 16pt default


# =============================================================================
# 3. DOCUMENT BUILDER HELPERS
# =============================================================================

def create_thai_document(
    margin_top_in: float = 1.0,
    margin_bottom_in: float = 1.0,
    margin_left_in: float = 1.0,
    margin_right_in: float = 1.0,
    font_name: str = DEFAULT_FONT,
    body_size_pt: float = DEFAULT_BODY_SIZE_PT
) -> docx.Document:
    """Create a new Document pre-configured with Thai default styles and margins."""
    doc = docx.Document()

    # 1. Base Normal style
    style_normal = doc.styles["Normal"]
    style_normal.font.name = font_name
    style_normal.font.size = Pt(body_size_pt)
    for attr in ["ascii", "hAnsi", "cs", "eastAsia"]:
        style_normal._element.rPr.rFonts.set(qn(f"w:{attr}"), font_name)

    # 2. Section Margins
    for s in doc.sections:
        s.top_margin = Inches(margin_top_in)
        s.bottom_margin = Inches(margin_bottom_in)
        s.left_margin = Inches(margin_left_in)
        s.right_margin = Inches(margin_right_in)

    return doc


def add_heading_1(
    doc: docx.Document,
    text: str,
    font_name: str = DEFAULT_FONT,
    size_pt: float = 18,
    color_rgb: RGBColor = RGBColor(27, 54, 93),
    space_before_pt: float = 14,
    space_after_pt: float = 4,
    clean: bool = True
):
    """Add Level 1 heading formatted for Thai documents."""
    p = doc.add_paragraph()
    set_paragraph_thai(
        p,
        alignment="left",
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
        keep_with_next=True
    )
    run_text = clean_thai_text(text) if clean else text
    run = p.add_run(run_text)
    set_run_font(run, name=font_name, size_pt=size_pt, bold=True, color_rgb=color_rgb)
    return p


def add_heading_2(
    doc: docx.Document,
    text: str,
    font_name: str = DEFAULT_FONT,
    size_pt: float = 16,
    color_rgb: RGBColor = RGBColor(0, 102, 153),
    space_before_pt: float = 10,
    space_after_pt: float = 3,
    clean: bool = True
):
    """Add Level 2 heading formatted for Thai documents."""
    p = doc.add_paragraph()
    set_paragraph_thai(
        p,
        alignment="left",
        space_before_pt=space_before_pt,
        space_after_pt=space_after_pt,
        keep_with_next=True
    )
    run_text = clean_thai_text(text) if clean else text
    run = p.add_run(run_text)
    set_run_font(run, name=font_name, size_pt=size_pt, bold=True, color_rgb=color_rgb)
    return p


def add_body_paragraph(
    doc: docx.Document,
    text: str,
    indent: bool = True,
    bold_prefix: str = None,
    font_name: str = DEFAULT_FONT,
    size_pt: float = DEFAULT_BODY_SIZE_PT,
    space_after_pt: float = 4,
    alignment: str = "thaiDistribute",
    clean: bool = True
):
    """
    Add a Thai body paragraph with optional first-line indent and bold prefix.
    Uses Thai Distributed alignment by default.
    """
    p = doc.add_paragraph()
    first_indent = 0.4 if indent else 0  # ~1 cm
    set_paragraph_thai(
        p,
        alignment=alignment,
        space_after_pt=space_after_pt,
        first_line_indent_in=first_indent
    )

    if bold_prefix:
        prefix_txt = clean_thai_text(bold_prefix) if clean else bold_prefix
        r_bold = p.add_run(prefix_txt.strip() + " ")
        set_run_font(r_bold, name=font_name, size_pt=size_pt, bold=True)

    body_txt = clean_thai_text(text) if clean else text
    r_body = p.add_run(body_txt)
    set_run_font(r_body, name=font_name, size_pt=size_pt)
    return p


def add_numbered_item(
    doc: docx.Document,
    num_str: str,
    text: str,
    bold_title: str = None,
    font_name: str = DEFAULT_FONT,
    size_pt: float = DEFAULT_BODY_SIZE_PT,
    left_indent_in: float = 0.45,
    first_line_indent_in: float = -0.25,
    num_color_rgb: RGBColor = RGBColor(27, 54, 93),
    alignment: str = "thaiDistribute",
    clean: bool = True
):
    """
    Add a numbered list item with proper hanging indent and Thai Distributed alignment.
    Eliminates ugly bullet/colon artifacts.
    """
    p = doc.add_paragraph()
    set_paragraph_thai(
        p,
        alignment=alignment,
        space_after_pt=3,
        left_indent_in=left_indent_in,
        first_line_indent_in=first_line_indent_in
    )

    num_txt = clean_thai_text(num_str) if clean else num_str
    r_num = p.add_run(num_txt.strip() + " ")
    set_run_font(r_num, name=font_name, size_pt=size_pt, bold=True, color_rgb=num_color_rgb)

    if bold_title:
        title_txt = clean_thai_text(bold_title) if clean else bold_title
        r_title = p.add_run(title_txt.strip() + " ")
        set_run_font(r_title, name=font_name, size_pt=size_pt, bold=True)

    body_txt = clean_thai_text(text) if clean else text
    r_body = p.add_run(body_txt)
    set_run_font(r_body, name=font_name, size_pt=size_pt)
    return p


# =============================================================================
# 4. TABLE STYLING & FORMATTING
# =============================================================================

def set_cell_background(cell, hex_color: str):
    """Set cell background shading."""
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_margins(cell, top: int = 110, bottom: int = 110, left: int = 130, right: int = 130):
    """Set cell internal padding in twips/dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("w:top", top), ("w:bottom", bottom), ("w:left", left), ("w:right", right)]:
        node = OxmlElement(m)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def style_table(
    table: docx.table.Table,
    col_widths: list = None,
    header_bg: str = "1B365D",
    alt_bg: str = "F4F6F9",
    border_color: str = "B0C4DE",
    font_name: str = DEFAULT_FONT,
    font_size_pt: float = 15,
    clean: bool = True
):
    """
    Style a table with professional headers, alternating row colors,
    clean horizontal borders, vertical centering, and Thai typography.
    """
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr

    # Clean border styling
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        f'<w:insideV w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    for i, row in enumerate(table.rows):
        is_header = (i == 0)
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if is_header:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for j, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=110, bottom=110, left=130, right=130)

            if is_header:
                set_cell_background(cell, header_bg)
            elif alt_bg and i % 2 == 1:
                set_cell_background(cell, alt_bg)

            if col_widths and j < len(col_widths):
                cell.width = Inches(col_widths[j])

            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.15
                set_paragraph_thai(
                    p,
                    alignment="center" if is_header else "thaiDistribute",
                    space_after_pt=0
                )
                for r in p.runs:
                    if clean:
                        r.text = clean_thai_text(r.text)
                    if is_header:
                        set_run_font(
                            r,
                            name=font_name,
                            size_pt=font_size_pt,
                            bold=True,
                            color_rgb=RGBColor(255, 255, 255)
                        )
                    else:
                        set_run_font(
                            r,
                            name=font_name,
                            size_pt=font_size_pt,
                            bold=False,
                            color_rgb=RGBColor(30, 30, 30)
                        )


# =============================================================================
# 5. REPAIR & INSPECTION OF EXISTING DOCUMENTS
# =============================================================================

def apply_thai_distribute_to_document(
    input_path: str,
    output_path: str = None,
    font_name: str = DEFAULT_FONT,
    to_arabic: bool = False,
    remove_colon: bool = True
) -> str:
    """
    Repair an existing .docx file by converting all justified or body paragraphs
    to Thai Distributed (thaiDistribute) and applying Complex Script font tagging.
    Fixes letter-stretching and word-gap bugs.
    """
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_thai_fixed{ext}"

    doc = docx.Document(input_path)

    for p in doc.paragraphs:
        # If paragraph is justified or has long text, set to thaiDistribute
        if p.alignment in [WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.THAI_JUSTIFY] or len(p.text) > 80:
            pPr = p._p.get_or_add_pPr()
            jc = pPr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                pPr.append(jc)
            jc.set(qn("w:val"), "thaiDistribute")
            set_paragraph_thai(p, alignment="thaiDistribute")

        for r in p.runs:
            if to_arabic or remove_colon:
                r.text = clean_thai_text(r.text, to_arabic=to_arabic, remove_colon=remove_colon)
            set_run_font(r, name=font_name, size_pt=r.font.size.pt if r.font.size else DEFAULT_BODY_SIZE_PT, bold=r.bold, italic=r.italic)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if to_arabic or remove_colon:
                            r.text = clean_thai_text(r.text, to_arabic=to_arabic, remove_colon=remove_colon)
                        set_run_font(r, name=font_name, size_pt=r.font.size.pt if r.font.size else DEFAULT_BODY_SIZE_PT, bold=r.bold, italic=r.italic)

    doc.save(output_path)
    return output_path


def inspect_document_alignments(docx_path: str) -> dict:
    """Inspect and count paragraph alignment types across an existing .docx file."""
    doc = docx.Document(docx_path)
    counts = {}
    for p in doc.paragraphs:
        al = str(p.alignment)
        counts[al] = counts.get(al, 0) + 1
    return counts

