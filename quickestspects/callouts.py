import os
from docx.enum.text import WD_BREAK
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION


def callout_section(doc, imgs_path):
    img_path = os.path.join(imgs_path, 'c08518669.png')
    img_path2 = os.path.join(imgs_path, 'c08518762.png')

    doc.add_picture(img_path, width=Inches(6))
    
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Left")
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    callout_table = doc.add_table(rows=4, cols=4)
    callout_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i in range(4):
        for j in range(4):
            cell = callout_table.cell(i, j)
            cell.text = str(i * 4 + j + 1)
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER


    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.add_picture(img_path2, width=Inches(6))

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Right")
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    callout_table2 = doc.add_table(rows=4, cols=4)
    callout_table2.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i in range(4):
        for j in range(4):
            cell = callout_table2.cell(i, j)
            cell.text = str(i * 4 + j + 1)
            cell.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break() 
    section = doc.sections[-1]
    section.start_type
    section.start_type = WD_SECTION.CONTINUOUS