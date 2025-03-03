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
    run.font.color.rgb = RGBColor.from_string(color)

def add_hyperlink(paragraph, text, url, font_name=DEFAULT_FONT, font_size=DEFAULT_SIZE, color=BLUE_COLOR):
    """Add a clickable hyperlink to a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    run = paragraph.add_run()
    run.text = text
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run._r.append(hyperlink)
    add_text_run(run, text, font_name, font_size, color=color, underline=True)

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
    validated_color = validate_color(color)
    fill = OxmlElement('v:fill')
    fill.set(qn('color2'), validated_color[1:])
    shape.append(fill)
    line.append(shape)
    paragraph._p.append(line)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(spacing_before)
    paragraph.paragraph_format.space_after = Pt(spacing_after)

def validate_color(color):
    """Validate and normalize a hexadecimal color code (e.g., '#000000' or '#FFF')."""
    if not isinstance(color, str):
        return "#000000"
    color = color.strip().lstrip('#')
    if not color:
        return "#000000"
    try:
        if len(color) not in (3, 6):
            return "#000000"
        int(color, 16)  # Test if it's a valid hex number
        if len(color) == 3:
            color = ''.join(c * 2 for c in color)
        return f"#{color}"
    except ValueError:
        return "#000000"

@app.route('/generate_cv', methods=['POST'])
def generate_cv():
    logger.debug("Received request for /generate_cv")
    try:
        # Validate and parse JSON
        if not request.is_json:
            logger.error("No JSON data received")
            return jsonify({"error": "Invalid JSON data"}), 400
        data = request.json
        logger.debug(f"JSON data received: {data}")

        # Ensure data is a dictionary (unwrap if it’s an array with a single object containing 'text')
        if isinstance(data, list) and len(data) == 1 and isinstance(data[0], dict) and 'text' in data[0]:
            try:
                # Parse the JSON string in 'text' to an object
                data = json.loads(data[0]['text'])
            except json.JSONDecodeError as e:
                logger.error(f"Failed to parse JSON text: {str(e)}")
                return jsonify({"error": "Invalid JSON text in request"}), 400
        elif not isinstance(data, dict):
            logger.error("Invalid JSON data structure: expected object or list with text field")
            return jsonify({"error": "Invalid JSON data structure"}), 400

        # Check for required fields in the schema
        required_fields = ['personalInformation', 'contactDetails', 'overview', 'workExperience', 'education', 'skills']
        if not all(field in data for field in required_fields):
            logger.error("Missing required fields in JSON data")
            return jsonify({"error": "Missing required fields in JSON data"}), 400

        # Create Word document
        doc = Document()

        # Set A4 page size and margins
        section = doc.sections[0]
        section.page_width = Cm(PAGE_WIDTH_CM)
        section.page_height = Cm(PAGE_HEIGHT_CM)
        section.top_margin = Cm(MARGINS_CM)
        section.bottom_margin = Cm(MARGINS_CM)
        section.left_margin = Cm(MARGINS_CM)
        section.right_margin = Cm(MARGINS_CM)

        # Header (Personal Information and Contact Details)
        # Name (large, bold, centered)
        name = data['personalInformation']['name'] if data['personalInformation']['name'] != "N/A" else "N/A"
        add_paragraph(doc, name, WD_ALIGN_PARAGRAPH.CENTER, DEFAULT_FONT, 20, bold=True)

        # Contact Details (centered, smaller font)
        contact_parts = []
        if data['contactDetails']['phone'] != "N/A":
            contact_parts.append(data['contactDetails']['phone'])
        if data['contactDetails']['email'] != "N/A":
            contact_parts.append(data['contactDetails']['email'])
        if data['contactDetails']['website'] != "N/A":
            contact_parts.append(data['contactDetails']['website'])
        if data['contactDetails']['location']['city'] != "N/A" and data['contactDetails']['location']['countryCode'] != "N/A":
            contact_parts.append(f"{data['contactDetails']['location']['city']}, {data['contactDetails']['location']['countryCode']}")
        contact_text = " | ".join(contact_parts) if contact_parts else "N/A"
        
        contact_para = add_paragraph(doc, "", WD_ALIGN_PARAGRAPH.CENTER, DEFAULT_FONT, 10)
        if data['contactDetails']['phone'] != "N/A":
            add_text_run(contact_para, data['contactDetails']['phone'] + " | ")
        if data['contactDetails']['email'] != "N/A":
            add_hyperlink(contact_para, data['contactDetails']['email'], f"mailto:{data['contactDetails']['email']}", DEFAULT_FONT, 10)
            add_text_run(contact_para, " | ")
        if data['contactDetails']['website'] != "N/A":
            add_hyperlink(contact_para, data['contactDetails']['website'], data['contactDetails']['website'], DEFAULT_FONT, 10)
            add_text_run(contact_para, " | ")
        if data['contactDetails']['location']['city'] != "N/A" and data['contactDetails']['location']['countryCode'] != "N/A":
            add_text_run(contact_para, f"{data['contactDetails']['location']['city']}, {data['contactDetails']['location']['countryCode']}")
        
        # Overview (Desired Role and Tagline, then bespoke sections)
        if data['overview']['desired_role'] != "N/A":
            add_paragraph(doc, data['overview']['desired_role'], WD_ALIGN_PARAGRAPH.CENTER, DEFAULT_FONT, 14, bold=True, spacing_before=12)
        if data['overview']['tagline'] != "N/A":
            add_paragraph(doc, data['overview']['tagline'], WD_ALIGN_PARAGRAPH.CENTER, DEFAULT_FONT, 12, italic=True, spacing_after=12)

        # Bespoke overview sections (e.g., Professional Overview, Career Highlights)
        for key, value in data['overview'].items():
            if key not in ['desired_role', 'tagline'] and isinstance(value, list) and value and value[0] != "N/A":
                add_paragraph(doc, key, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
                if key == 'Career Highlights':
                    add_bullet_list(doc, value, DEFAULT_FONT, DEFAULT_SIZE)
                else:
                    for line in value:
                        add_paragraph(doc, line, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE, spacing_after=6)

        # Work Experience
        if data['workExperience']:
            add_paragraph(doc, "Career Experience", WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
            for exp in data['workExperience']:
                for pos in exp['position']:
                    role_text = f"{pos['jobTitle']} - {exp['organisation']} | {exp['about_the_organisation']}, {exp['location']} ({pos['startDate']} – {pos['endDate'] if pos['endDate'] != 'N/A' else 'Present'})"
                    add_paragraph(doc, role_text, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 12, bold=True, spacing_after=6)
                    if 'plainText' in pos['details'] and pos['details']['plainText'] != "N/A":
                        if isinstance(pos['details']['plainText'], list):
                            add_bullet_list(doc, pos['details']['plainText'], DEFAULT_FONT, DEFAULT_SIZE)
                        else:
                            add_paragraph(doc, pos['details']['plainText'], WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE)
                    for key, value in pos['details'].items():
                        if key not in ['plainText'] and isinstance(value, list) and value and value[0] != "N/A":
                            add_paragraph(doc, key, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 11, bold=True, spacing_before=6)
                            add_bullet_list(doc, value, DEFAULT_FONT, DEFAULT_SIZE, indent=Cm(1.27))

        # Education
        if data['education']:
            add_paragraph(doc, "Education", WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
            for edu in data['education']:
                edu_text = f"{edu['studyType']} in {edu['area']} - {edu['institution']}, {edu['location'] if edu['location'] != 'N/A' else ''} ({edu['startDate']} – {edu['endDate'] if edu['endDate'] != 'N/A' else 'Present'})"
                add_paragraph(doc, edu_text, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE, spacing_after=6)
                if edu['score'] != "N/A":
                    add_paragraph(doc, f"Score: {edu['score']}", WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE)

        # Skills (three-column format for Key Skills & Expertise)
        if data['skills']:
            add_paragraph(doc, "Key Skills & Expertise", WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
            columns = [
                {"text": "\n".join([s['name'] for s in data['skills'][:len(data['skills'])//3]]), "font": {"name": DEFAULT_FONT, "size": DEFAULT_SIZE, "color": BLACK_COLOR}},
                {"text": "\n".join([s['name'] for s in data['skills'][len(data['skills'])//3:2*len(data['skills'])//3]]), "font": {"name": DEFAULT_FONT, "size": DEFAULT_SIZE, "color": BLACK_COLOR}},
                {"text": "\n".join([s['name'] for s in data['skills'][2*len(data['skills'])//3:]]), "font": {"name": DEFAULT_FONT, "size": DEFAULT_SIZE, "color": BLACK_COLOR}}
            ]
            create_three_column_table(doc, columns)

        # Associations, Publications, Projects, Volunteer, Interests, Patents, Awards, Certificates, Languages (if present)
        for section in ['associations', 'publications', 'projects', 'volunteer', 'interests', 'patents', 'awards', 'certificates', 'languages']:
            if data.get(section) and data[section] and (isinstance(data[section], list) and data[section][0] != "N/A" or data[section] != "N/A"):
                title = section.capitalize().replace('_', ' ')
                add_paragraph(doc, title, WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
                if isinstance(data[section], list):
                    add_bullet_list(doc, [item if isinstance(item, str) else str(item) for item in data[section] if item != "N/A"], DEFAULT_FONT, DEFAULT_SIZE)
                else:
                    add_paragraph(doc, data[section], WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE)

        # References
        if data['references']:
            add_paragraph(doc, "References", WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, 14, bold=True, spacing_before=12)
            for ref in data['references']:
                if ref['name'] != "N/A" and ref['reference'] != "N/A":
                    add_paragraph(doc, ref['name'], WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE, bold=True)
                    add_paragraph(doc, ref['reference'], WD_ALIGN_PARAGRAPH.LEFT, DEFAULT_FONT, DEFAULT_SIZE, spacing_after=6)

        # Save document to temporary directory (Render-safe)
        output_filename = f"Vaughn_Marles_CV_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        output_path = os.path.join('/tmp', output_filename)
        os.makedirs('/tmp', exist_ok=True)
        doc.save(output_path)
        logger.debug(f"Word document generated and saved at {output_path}")
        return send_file(output_path, as_attachment=True, download_name=output_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        logger.error(f"Error in /generate_cv: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    # Use Render's PORT (no fallback) for cloud deployment
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
