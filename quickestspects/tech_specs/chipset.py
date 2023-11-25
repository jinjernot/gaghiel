from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def chipset_section(doc, txt_file, df):

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("CHIPSET")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write("<h1><b>CHIPSET</h1></b>\n")

    chipset = df.iloc[90, 6]
    paragraph.add_run(chipset)

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{chipset}</p>\n")

    insertHR(doc.add_paragraph(), thickness=3)

    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')
