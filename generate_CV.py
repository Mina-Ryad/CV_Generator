from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
import json

def add_heading(document, text, level=1):
    heading = document.add_heading(level=level)
    run = heading.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    heading.paragraph_format.space_before = Pt(6)
    heading.paragraph_format.space_after = Pt(3)

def add_name(document, name):
    name_paragraph = document.add_paragraph(name)
    name_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    name_paragraph.style = 'Heading 1'
    name_paragraph.style.font.size = Pt(20)
    reduce_spacing(name_paragraph)

def add_contact_info(document, contact_info):
    add_heading(document, 'Contact Information')
    for key, value in contact_info.items():
        p = document.add_paragraph()
        p.add_run(f"{key}: ").bold = True
        p.add_run(value)
        reduce_spacing(p)

def add_section(document, title, items):
    add_heading(document, title)
    for item in items:
        p = document.add_paragraph(item, style='ListBullet')  # Removed manual bullet
        p.style.font.size = Pt(10)
        reduce_spacing(p)

def add_experience(document, experiences):
    add_heading(document, 'Work Experience')
    for exp in experiences:
        # Add job title and company in bold without bullet points
        title = f"{exp['title']} at {exp['company']} ({exp['dates']})"
        p = document.add_paragraph()
        p.add_run(title).bold = True
        reduce_spacing(p)
        
        # Add achievements as bullet points
        for detail in exp['achievements']:
            d = document.add_paragraph(detail, style='ListBullet')  # Removed manual bullet
            reduce_spacing(d)

def reduce_spacing(paragraph, line_spacing_rule=WD_LINE_SPACING.EXACTLY, line_spacing=Pt(12), space_after=Pt(0), space_before=Pt(1)):
    paragraph.paragraph_format.line_spacing_rule = line_spacing_rule
    paragraph.paragraph_format.line_spacing = line_spacing
    paragraph.paragraph_format.space_after = space_after
    paragraph.paragraph_format.space_before = space_before

def read_cv_data(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

def generate_cv(data):
    document = Document()
    document.styles['Heading 1'].font.size = Pt(16)
    document.styles['Heading 2'].font.size = Pt(12)
    
    for section in document.sections:
        section.top_margin = Pt(40)
        section.bottom_margin = Pt(40)
        section.left_margin = Pt(40)
        section.right_margin = Pt(40)

    add_name(document, data['name'])
    add_contact_info(document, data['contact_info'])
    add_section(document, 'Professional Summary', [data['summary']])
    add_experience(document, data['experience'])
    add_section(document, 'Education', data['education'])
    add_section(document, 'Skills', data['skills'])
    add_section(document, 'Certifications', data['certifications'])
    add_section(document, 'Projects', data['projects'])
    add_section(document, 'Interests', data['interests'])

    document.save('ATS_Compliant_CV.docx')

if __name__ == "__main__":
    cv_data = read_cv_data("cv_data.json")
    generate_cv(cv_data)
