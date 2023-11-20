
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def product_name_section(doc, txt_file, prod_name):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("PRODUCT NAME")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    with open(txt_file, 'a') as txt:
        txt.write("<h1>Technical Specifications</h1>\n")
        txt.write("<h1>PRODUCT NAME</h1>\n")
        txt.write(f"<p>{prod_name}</p>\n")
    
    prod_name_paragraph = doc.add_paragraph(prod_name)

    insertHR(doc.add_paragraph(), thickness=3)

    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')