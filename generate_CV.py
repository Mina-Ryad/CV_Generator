from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import json

def set_document_formatting(document):
    style = document.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)
    style.paragraph_format.line_spacing = 1.15
    for section in document.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

def add_section_heading(document, title, level=1):
    heading = document.add_heading(level=level)
    run = heading.add_run(title)
    run.bold = True
    run.font.size = Pt(12)
    heading.paragraph_format.space_before = Pt(6)
    heading.paragraph_format.space_after = Pt(4)
    return heading

def add_contact_info(document, data):
    p = document.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run(data["name"].upper())
    run.bold = True
    run.font.size = Pt(14)
    p.add_run("\n")
    run = p.add_run(data["title"])
    run.italic = True
    run.font.size = Pt(11)
    p.add_run("\n")
    contact = data["contact"]
    location_phone = f"{contact['location']} | {contact['phone']} | {contact['email']}"
    run = p.add_run(location_phone)
    run.font.size = Pt(9)
    p.add_run("\n")
    linkedin_run = p.add_run(f"LinkedIn: {contact['linkedin']}")
    linkedin_run.font.color.rgb = RGBColor(0, 0, 255)
    linkedin_run.underline = True
    p.add_run("\n")
    run = p.add_run(f"Languages: {contact['languages']}")
    run.font.size = Pt(9)
    p.paragraph_format.space_after = Pt(8)

def add_bullet_list(document, items):
    for item in items:
        p = document.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(item.replace('--', '–'))  # Replace hyphens with en-dashes
        run.font.size = Pt(10)

def generate_cv(data):
    doc = Document()
    set_document_formatting(doc)
    add_contact_info(doc, data)
    add_section_heading(doc, "SUMMARY")
    doc.add_paragraph(data["sections"]["summary"]["content"]).paragraph_format.space_after = Pt(6)
    add_section_heading(doc, "EXPERIENCE")
    for exp in data["sections"]["experience"]:
        p = doc.add_paragraph()
        p.add_run(f"{exp['company']} | {exp['location']}\n").bold = True
        p.add_run(f"{exp['position']} | {exp['dates']}").italic = True
        p.paragraph_format.space_after = Pt(4)
        add_bullet_list(doc, exp["highlights"])
    add_section_heading(doc, "EDUCATION")
    for edu in data["sections"]["education"]:
        text = f"{edu['degree']} - {edu['institution']} | {edu['dates']}"
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Inches(0.25)
        p.add_run(text).font.size = Pt(10)
    add_section_heading(doc, "TECHNICAL SKILLS")
    add_bullet_list(doc, data["sections"]["technical_skills"])
    add_section_heading(doc, "CERTIFICATIONS & TRAINING")
    add_bullet_list(doc, data["sections"]["certifications"])
    add_section_heading(doc, "PROJECTS")
    add_bullet_list(doc, data["sections"]["projects"])
    add_section_heading(doc, "ADDITIONAL INFORMATION")
    add_bullet_list(doc, data["sections"]["additional"]["items"])
    doc.save("Mina_Ryad_CV_Final.docx")

if __name__ == "__main__":
    with open('cv_data.json') as f:
        data = json.load(f)
    generate_cv(data)