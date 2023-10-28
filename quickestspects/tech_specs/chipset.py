
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def chipset_section(doc, df):

    chipset_paragraph = doc.add_paragraph()
    run = chipset_paragraph.add_run("CHIPSET")
    run.font.size = Pt(12)
    run.bold = True
    chipset_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    chipset_paragraph.add_run().add_break()

    chipset = df.iloc[90, 6]
    chipset_paragraph.add_run(chipset)
    insertHR(doc.add_paragraph(), thickness=3)
