from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
import json

def add_heading(document, text, level=1):
    # Add a heading to the document with the specified text and level.
    heading = document.add_heading(level=level)
    run = heading.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    # Reduce spacing in the paragraph
    paragraph_format = heading.paragraph_format
    paragraph_format.space_before = Pt(8)
    paragraph_format.space_after = Pt(1)

def add_name(document, name):
    # Add the name to the document.
    name_paragraph = document.add_paragraph(name)
    name_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    name_paragraph.style = 'Heading 1'
    name_paragraph.style.font.size = Pt(24)
    reduce_spacing(name_paragraph)

def add_personal_info(document, personal_info):
    # Add personal information section to the document.
    add_heading(document, 'Personal Information')
    for key, value in personal_info.items():
        key_paragraph = document.add_paragraph()
        key_run = key_paragraph.add_run(key + ': ')
        key_run.bold = True
        key_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        key_paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(4000))
        value_run = key_paragraph.add_run(value)
        if key.lower() == 'linkedin':
            value_run.font.color.rgb = RGBColor(0, 0, 255)  # Blue color
            value_run.underline = True  # Underline style
        reduce_spacing(key_paragraph)

def add_section(document, title, dates, content, bullet=False, font_size=None):
    # Add a section with the specified title, dates, and content to the document.
    document.add_heading(title, level=2)
    if dates:
        paragraph = document.add_paragraph()
        run = paragraph.add_run(f"{dates}")
        run.italic = True
        reduce_spacing(paragraph, space_after=Pt(3))
    if bullet:
        paragraphs = content.split('\n')
        for para in paragraphs:
            if para.strip():
                p = document.add_paragraph(para, style='ListBullet')
                if font_size:
                    p.style.font.size = Pt(font_size)
                reduce_spacing(p)
    else:
        paragraph = document.add_paragraph(content)
        if font_size:
            paragraph.style.font.size = Pt(font_size)
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        reduce_spacing(paragraph)

def add_summary(document, summary):
    # Add a summary to the document.
    add_heading(document, 'Summary')
    summary_paragraph = document.add_paragraph(summary)
    summary_paragraph.style.font.size = Pt(10)
    reduce_spacing(summary_paragraph)

def add_experience(document, experiences):
    # Add experience sections to the document.
    add_heading(document, 'Experience')
    for exp in experiences:
        exp_title = f"{exp['title']}, {exp['company']}"
        exp_dates = exp['dates']
        exp_details = exp['description']
        add_section(document, exp_title, exp_dates, exp_details, bullet=True, font_size=10)

def add_education(document, education):
    # Add education sections to the document.
    add_heading(document, 'Education')
    for edu in education:
        degree = edu['degree']
        school = edu['school']
        dates = edu['dates']
        # Create a new paragraph for each education entry
        p = document.add_paragraph()
        # Add bullet point
        p.add_run("\u2022 ").bold = True
        # Add degree in bold
        p.add_run(f" {degree}").bold = True
        # Add school
        p.add_run(f", {school}")
        # Add dates in italic without the comma
        p.add_run(f" {dates}").italic = True

def add_courses(document, courses):
    # Add courses sections to the document.
    add_heading(document, 'Courses')
    for course in courses:
        CoursesParagraph = document.add_paragraph(course, style='ListBullet')
        CoursesParagraph.style.font.size = Pt(10)
        reduce_spacing(CoursesParagraph)

def add_technical_skills(document, skills):
    # Add technical skills sections to the document.
    add_heading(document, 'Technical Skills')
    for skill in skills:
        CoursesParagraph = document.add_paragraph(skill, style='ListBullet')
        CoursesParagraph.style.font.size = Pt(10)
        reduce_spacing(CoursesParagraph)

def add_projects(document, projects):
    # Add projects sections to the document.
    add_heading(document, 'Projects')
    for project in projects:
        ProjectsParagraph = document.add_paragraph(project, style='ListBullet')
        ProjectsParagraph.style.font.size = Pt(10)
        reduce_spacing(ProjectsParagraph)

def add_hobbies(document, hobbies):
    # Add hobbies and interests sections to the document.
    add_heading(document, 'Hobbies and Interests')
    for hobby in hobbies:
        HobbiesParagraph = document.add_paragraph(hobby, style='ListBullet')
        HobbiesParagraph.style.font.size = Pt(10)
        reduce_spacing(HobbiesParagraph)

def reduce_spacing(paragraph, line_spacing_rule=WD_LINE_SPACING.EXACTLY, line_spacing=Pt(12), space_after=Pt(0), space_before=Pt(1)):
    # Reduce spacing in the paragraph.
    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing_rule = line_spacing_rule
    paragraph_format.line_spacing = line_spacing
    paragraph_format.space_after = space_after
    paragraph_format.space_before = space_before

def read_cv_data_from_json(file_path):
    # Read CV data from a JSON file.
    with open(file_path, 'r', encoding='utf-8') as file:
        cv_data = json.load(file)
    return cv_data

def generate_cv(data):
    # Generate a DOCX version of a CV document based on the provided data.
    document = Document()

    # Increase the font size of main headers
    document.styles['Heading 1'].font.size = Pt(16)
    document.styles['Heading 2'].font.size = Pt(11)

    # Decrease page borders
    sections = document.sections
    for section in sections:
        section.top_margin = Pt(50)
        section.bottom_margin = Pt(50)
        section.left_margin = Pt(50)
        section.right_margin = Pt(50)

    # Add name at the top
    add_name(document, data['name'])

    add_personal_info(document, data['personal_info'])
    add_summary(document, data['summary'])
    add_experience(document, data['experience'])
    add_education(document, data['education'])
    add_courses(document, data['courses'])
    add_technical_skills(document, data['technical_skills'])
    add_projects(document, data['projects'])
    add_hobbies(document, data['hobbies'])

    # Save the DOCX version
    docx_file = 'Mina_Ryad_CV.docx'
    document.save(docx_file)

if __name__ == "__main__":
    cv_data = read_cv_data_from_json("cv_data.json")
    generate_cv(cv_data)
