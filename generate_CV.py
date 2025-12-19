"""
CV Generator Script (Final Polished Version)
Author: Mina Ryad
Description: Generates a 2-page CV with correct bullet point spacing (Fixes the wide gap issue).
"""

import json
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --------------------------
# Configuration Constants
# --------------------------
FONT_NAME = 'Arial'
FONT_SIZE_NORMAL = Pt(10)        # Standard readable size
FONT_SIZE_HEADER = Pt(16)        # Name size
FONT_SIZE_SUBHEADER = Pt(10)     # Job Titles
FONT_SIZE_SECTION_TITLE = Pt(11) # Section Headers
LINE_SPACING = 1.05              # Tight but readable
MARGINS = Inches(0.5)            # Narrow margins to fit 2 pages
PAGE_WIDTH_TILES = Inches(7.5)   # 8.5" - 0.5" - 0.5" = 7.5"

SECTION_TITLES = {
    'summary': 'SUMMARY',
    'experience': 'PROFESSIONAL EXPERIENCE',
    'education': 'EDUCATION',
    'skills': 'TECHNICAL SKILLS',
    'certifications': 'CERTIFICATIONS & TRAINING',
    'projects': 'KEY PROJECTS',
    'additional': 'ADDITIONAL INFORMATION'
}

def set_document_formatting(doc):
    """Sets global document formatting."""
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = FONT_SIZE_NORMAL
    style.paragraph_format.line_spacing = LINE_SPACING
    
    for section in doc.sections:
        section.top_margin = MARGINS
        section.bottom_margin = MARGINS
        section.left_margin = MARGINS
        section.right_margin = MARGINS

def add_bottom_border(paragraph):
    """Adds a real bottom border to a paragraph using OXML."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')      
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

def create_section_header(doc, title):
    """Creates a standardized section header with a bottom border."""
    heading = doc.add_heading(level=1)
    run = heading.add_run(title)
    run.bold = True
    run.font.size = FONT_SIZE_SECTION_TITLE
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = FONT_NAME
    
    add_bottom_border(heading)
    
    # COMPACT SPACING
    heading.paragraph_format.space_before = Pt(12) 
    heading.paragraph_format.space_after = Pt(4)   
    return heading

def add_contact_info(doc, header_data):
    """Adds contact information section."""
    contact_para = doc.add_paragraph()
    contact_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    contact_para.paragraph_format.space_after = Pt(8) 
    
    # Name
    name_run = contact_para.add_run(header_data['name'].upper())
    name_run.bold = True
    name_run.font.size = FONT_SIZE_HEADER
    contact_para.add_run('\n')
    
    # Title
    title_run = contact_para.add_run(header_data['title'])
    title_run.font.size = Pt(10)
    contact_para.add_run('\n')
    
    # Details
    details = [
        header_data['contact']['location'],
        header_data['contact']['phone'],
        header_data['contact']['email'],
        header_data['contact']['linkedin']
    ]
    details = [d for d in details if d]
    
    separator = "  |  "
    for i, detail in enumerate(details):
        if "http" in detail or "www" in detail:
            run = contact_para.add_run(detail)
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.underline = True
        else:
            contact_para.add_run(detail)
        
        if i < len(details) - 1:
            contact_para.add_run(separator)

def add_two_column_line(doc, left_text, right_text, is_bold=True, is_italic=False):
    """Line with text on left and right (Company ... Date)."""
    para = doc.add_paragraph()
    tab_stops = para.paragraph_format.tab_stops
    tab_stops.add_tab_stop(PAGE_WIDTH_TILES, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)
    
    run_left = para.add_run(left_text)
    run_left.bold = is_bold
    run_left.font.size = FONT_SIZE_SUBHEADER
    
    para.add_run('\t')
    
    run_right = para.add_run(right_text)
    run_right.bold = is_bold 
    run_right.italic = is_italic
    run_right.font.size = FONT_SIZE_SUBHEADER
    
    para.paragraph_format.space_after = Pt(0)

def add_experience(doc, experience_items):
    create_section_header(doc, SECTION_TITLES['experience'])
    
    for job in experience_items:
        add_two_column_line(doc, job['company'], job['dates'], is_bold=True)
        
        # Position
        pos_para = doc.add_paragraph()
        pos_run = pos_para.add_run(job['position'])
        pos_run.italic = True
        pos_para.paragraph_format.space_after = Pt(2) 
        
        # Details
        for detail in job['details']:
            para = doc.add_paragraph(style='List Bullet')
            # STANDARD INDENT: Small gap (0.25")
            para.paragraph_format.left_indent = Inches(0.25)
            para.paragraph_format.space_after = Pt(0)
            para.add_run(detail)
        
        # Space between jobs
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_education(doc, education_items):
    create_section_header(doc, SECTION_TITLES['education'])
    for edu in education_items:
        add_two_column_line(doc, edu['degree'], edu['dates'], is_bold=True)
        para = doc.add_paragraph()
        run = para.add_run(edu['institution'])
        run.italic = True
        para.paragraph_format.space_after = Pt(6)

def add_smart_list_section(doc, section_key, items):
    """
    Adds a list section.
    FIX: Uses a standard tight indent (0.3") instead of a deep hanging indent (1.4")
    to prevent large gaps between the bullet and the text.
    """
    create_section_header(doc, SECTION_TITLES[section_key])
    for item in items:
        para = doc.add_paragraph(style='List Bullet')
        
        # FIX: Reduced indent to 0.3 inches
        # This keeps wrapping lines aligned, but brings text close to the bullet.
        para.paragraph_format.left_indent = Inches(0.3) 
        para.paragraph_format.first_line_indent = Inches(-0.3)
        para.paragraph_format.space_after = Pt(1) 
        
        if ":" in str(item):
            title, description = str(item).split(":", 1)
            run_title = para.add_run(title + ":")
            run_title.bold = True
            para.add_run(" " + description.strip())
        else:
            para.add_run(str(item))

def generate_cv(json_file):
    if not os.path.exists(json_file):
        print(f"Error: File '{json_file}' not found.")
        return

    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    doc = Document()
    set_document_formatting(doc)
    
    add_contact_info(doc, data['header'])
    
    content = data['sections']
    
    create_section_header(doc, SECTION_TITLES['summary'])
    summary_para = doc.add_paragraph(content['summary'])
    summary_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    
    add_experience(doc, content['experience'])
    add_education(doc, content['education'])
    
    if 'skills' in content:
        add_smart_list_section(doc, 'skills', content['skills'])

    if 'certifications' in content:
        create_section_header(doc, SECTION_TITLES['certifications'])
        for item in content['certifications']:
            p = doc.add_paragraph(style='List Bullet')
            # Standard indent for certs too
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.space_after = Pt(0)
            p.add_run(item)
    
    if 'projects' in content:
        add_smart_list_section(doc, 'projects', content['projects'])
        
    if 'additional' in content:
        add_smart_list_section(doc, 'additional', content['additional'])
    
    output_name = f"{data['header']['name'].replace(' ', '_')}_CV.docx"
    doc.save(output_name)
    print(f"CV successfully generated: {output_name}")

if __name__ == "__main__":
    generate_cv('cv_data.json')