from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import json
import os
from datetime import datetime
import logging

# Configure logging for debugging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Hard-coded formatting constants
DEFAULT_FONT = 'Arial'
DEFAULT_SIZE = 11  # pt
BLACK_COLOR = '#000000'
BLUE_COLOR = '#0000FF'  # For clickable links (email, URLs)
LINE_SPACING = 1.15  # Line spacing multiplier
MARGINS_CM = 2.54  # A4 margins in cm
PAGE_WIDTH_CM = 21.0  # A4 width in cm
PAGE_HEIGHT_CM = 29.7  # A4 height in cm

def add_text_run(paragraph, text, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, bold=False, italic=False, 
                 underline=False, color=BLACK_COLOR):
    """Add formatted text to a paragraph."""
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    try:
        run.font.color.rgb = RGBColor.from_string(validate_color(color))
    except ValueError as e:
        logger.warning(f"Invalid color {color} in add_text_run, defaulting to black: {str(e)}")
        run.font.color.rgb = RGBColor.from_string(BLACK_COLOR)

def add_hyperlink(paragraph, text, url, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, color=BLUE_COLOR):
    """Add a clickable hyperlink to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    run = paragraph.add_run()
    run.text = text
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run._r.append(hyperlink)
    try:
        add_text_run(run, text, font_name, font_size, color=validate_color(color), underline=True)
    except ValueError as e:
        logger.warning(f"Invalid color {color} in add_hyperlink, defaulting to blue: {str(e)}")
        add_text_run(run, text, font_name, font_size, color=BLUE_COLOR, underline=True)

def add_paragraph(doc, text, alignment=WD_ALIGN_PARAGRAPH.LEFT, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, 
                  bold=False, italic=False, spacing_before=0, spacing_after=6, line_spacing=LINE_SPACING):
    """Add a paragraph with custom formatting."""
    paragraph = doc.add_paragraph()
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(spacing_before)
    paragraph.paragraph_format.space_after = Pt(spacing_after)
    paragraph.paragraph_format.line_spacing = line_spacing
    add_text_run(paragraph, text, font_name, font_size, bold, italic)
    return paragraph

def add_bullet_list(doc, items, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, indent=Cm(0.63), 
                    spacing_before=0, spacing_after=6, line_spacing=LINE_SPACING):
    """Add a bulleted list to the document."""
    for item in items:
        if item and item != "N/A":
            paragraph = doc.add_paragraph(style='List Bullet')
            paragraph.paragraph_format.left_indent = indent
            paragraph.paragraph_format.space_before = Pt(spacing_before)
            paragraph.paragraph_format.space_after = Pt(spacing_after)
            paragraph.paragraph_format.line_spacing = line_spacing
            add_text_run(paragraph, item, font_name, font_size)

def create_three_column_table(doc, columns, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, 
                             spacing_before=0, spacing_after=6, line_spacing=LINE_SPACING):
    """Create a three-column table for skills or similar sections."""
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    column_width = Cm((PAGE_WIDTH_CM - (2 * MARGINS_CM)) / 3)  # Equal width for each column on A4, accounting for margins
    for col in table.columns:
        col.width = column_width

    for i, col_data in enumerate(columns):
        if col_data and col_data['text'] != "N/A":
            cell = table.cell(0, i)
            for text in col_data['text'].split('\n'):
                if text.strip():
                    paragraph = cell.add_paragraph()
                    add_text_run(paragraph, text.strip(), font_name, font_size)
                    paragraph.paragraph_format.space_before = Pt(spacing_before)
                    paragraph.paragraph_format.space_after = Pt(spacing_after)
                    paragraph.paragraph_format.line_spacing = line_spacing

    table.alignment = WD_ALIGN_PARAGRAPH.LEFT

def add_line(doc, width_percent='100%', height_pt=1, color=BLACK_COLOR, 
             spacing_before=12, spacing_after=6):
    """Add a horizontal line to the document."""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    line = OxmlElement('w:pict')
    shape = OxmlElement('v:rect')
    shape.set(qn('v:style'), 'solid')
    width_value = float(width_percent.strip('%')) / 100 * (PAGE_WIDTH_CM - (2 * MARGINS_CM)) * 36000 / 2.54  # Convert % to twips
    shape.set(qn('v:width'), str(int(width_value)))  # Width in EMUs (1 cm = 914400 EMUs)
    shape.set(qn('v:height'), str(int(height_pt * 12700)))  # Height in EMUs (1 pt = 12700 EMUs)
    try:
        validated_color = validate_color(color)
        fill = OxmlElement('v:fill')
        fill.set(qn('color2'), validated_color[1:])
        shape.append(fill)
    except ValueError as e:
        logger.warning(f"Invalid color {color} in add_line, defaulting to black: {str(e)}")
        fill = OxmlElement('v:fill')
        fill.set(qn('color2'), BLACK_COLOR[1:])
        shape.append(fill)
    line.append(shape)
    paragraph._p.append(line)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(spacing_before)
    paragraph.paragraph_format.space_after = Pt(spacing_after)

def validate_color(color):
    """Validate and normalize a hexadecimal color code, defaulting to black if invalid."""
    if not isinstance(color, str):
        logger.warning(f"Invalid color type for value: {color}, defaulting to #000000")
        return BLACK_COLOR
    color = color.strip().lstrip('#')
    if not color:
        logger.warning(f"Empty color value, defaulting to #000000")
        return BLACK_COLOR
    try:
        # Ensure it's a valid hex color (3 or 6 chars)
        if len(color) not in (3, 6):
            logger.warning(f"Invalid hex color length: {color} (must be 3 or 6 chars after '#'), defaulting to #000000")
            return BLACK_COLOR
        int(color, 16)  # Test if it's a valid hex number
        # Pad to 6 chars if 3 chars (e.g., 'FFF' -> 'FFFFFF')
        if len(color) == 3:
            color = ''.join(c * 2 for c in color)
        return f"#{color}"
    except ValueError as e:
        logger.warning(f"Invalid hex color: {color} - {str(e)}, defaulting to #000000")
        return BLACK_COLOR

def validate_json_colors(data):
    """Recursively validate and fix color values in the JSON to prevent invalid hex codes like '#0'."""
    if isinstance(data, dict):
        for key in list(data.keys()):
