
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def product_name_section(doc, prod_name):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("PRODUCT NAME")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    prod_name_paragraph = doc.add_paragraph(prod_name)

    insertHR(doc.add_paragraph(), thickness=3)