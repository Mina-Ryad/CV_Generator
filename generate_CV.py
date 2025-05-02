"""
CV Generator Script
Author: Mina Ryad
Description: Generates a professional CV from a JSON file input
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import json

# --------------------------
# Configuration Constants
# --------------------------
FONT_NAME = 'Arial'
FONT_SIZE_NORMAL = Pt(10)
FONT_SIZE_HEADER = Pt(14)
FONT_SIZE_SECTION_TITLE = Pt(12)
LINE_SPACING = 1.15
MARGINS = Inches(0.75)
SECTION_SPACING = Pt(6)

# --------------------------
# Section Display Titles
# --------------------------
SECTION_TITLES = {
    'summary': 'Professional Summary',
    'experience': 'Professional Experience',
    'education': 'Education',
    'skills': 'Technical Skills',
    'certifications': 'Certifications & Training',
    'projects': 'Projects',
    'additional': 'Additional Information'
}

def set_document_formatting(doc):
    """
    Sets global document formatting
    Args:
        doc: Word document object
    """
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = FONT_SIZE_NORMAL
    style.paragraph_format.line_spacing = LINE_SPACING
    
    for section in doc.sections:
        section.top_margin = MARGINS
        section.bottom_margin = MARGINS
        section.left_margin = MARGINS
        section.right_margin = MARGINS

def create_section_header(doc, title):
    """
    Creates a standardized section header
    Args:
        doc: Word document object
        title: Section title text
    Returns:
        Created paragraph object
    """
    heading = doc.add_heading(level=1)
    run = heading.add_run(title.upper())
    run.bold = True
    run.font.size = FONT_SIZE_SECTION_TITLE
    heading.paragraph_format.space_before = SECTION_SPACING
    heading.paragraph_format.space_after = Pt(4)
    return heading

def add_contact_info(doc, header_data):
    """
    Adds contact information section
    Args:
        doc: Word document object
        header_data: Dictionary containing contact information
    """
    contact_para = doc.add_paragraph()
    contact_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Name
    name_run = contact_para.add_run(header_data['name'].upper())
    name_run.bold = True
    name_run.font.size = FONT_SIZE_HEADER
    contact_para.add_run('\n')
    
    # Title
    title_run = contact_para.add_run(header_data['title'])
    title_run.italic = True
    contact_para.add_run('\n')
    
    # Contact Details
    details = [
        header_data['contact']['location'],
        header_data['contact']['phone'],
        header_data['contact']['email'],
        f"LinkedIn: {header_data['contact']['linkedin']}",
        f"Languages: {header_data['contact']['languages']}"
    ]
    
    for detail in details:
        run = contact_para.add_run(detail + '\n')
        if 'LinkedIn' in detail:
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.underline = True

def add_generic_section(doc, section_key, items):
    """
    Adds a bullet-point list section
    Args:
        doc: Word document object
        section_key: Key from SECTION_TITLES
        items: List of content items
    """
    create_section_header(doc, SECTION_TITLES[section_key])
    for item in items:
        para = doc.add_paragraph(style='List Bullet')
        para.paragraph_format.left_indent = Inches(0.25)
        para.add_run(str(item))

def add_experience(doc, experience_items):
    """
    Adds professional experience section
    Args:
        doc: Word document object
        experience_items: List of job experiences
    """
    create_section_header(doc, SECTION_TITLES['experience'])
    
    for job in experience_items:
        # Company/Position line
        company_para = doc.add_paragraph()
        company_para.add_run(f"{job['company']} | ").bold = True
        company_para.add_run(f"{job['position']} | {job['dates']}")
        
        # Job details
        for detail in job['details']:
            para = doc.add_paragraph(style='List Bullet')
            para.paragraph_format.left_indent = Inches(0.25)
            para.add_run(detail)

def generate_cv(json_file):
    """
    Main function to generate CV from JSON input
    Args:
        json_file: Path to JSON input file
    """
    with open(json_file) as f:
        data = json.load(f)
    
    doc = Document()
    set_document_formatting(doc)
    
    # Add header section
    add_contact_info(doc, data['header'])
    
    # Add main content sections
    content = data['sections']
    
    # Professional Summary
    create_section_header(doc, SECTION_TITLES['summary'])
    doc.add_paragraph(content['summary'])
    
    # Experience
    add_experience(doc, content['experience'])
    
    # Education
    add_generic_section(doc, 'education', [
        f"{edu['degree']} - {edu['institution']} | {edu['dates']}" 
        for edu in content['education']
    ])
    
    # Other sections
    for section in ['skills', 'certifications', 'projects', 'additional']:
        if section in content:
            add_generic_section(doc, section, content[section])
    
    # Save document
    output_name = f"{data['header']['name'].replace(' ', '_')}_CV.docx"
    doc.save(output_name)
    print(f"CV successfully generated: {output_name}")

if __name__ == "__main__":
    generate_cv('cv_data.json')