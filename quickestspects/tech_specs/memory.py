from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def memory_section(doc, txt_file, df):

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("MEMORY")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write("<h1><b>MEMORY</h1></b>\n")

    maximum_memory_subtitle = df.iloc[235, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(maximum_memory_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    maximum_memory = df.iloc[236, 6]
    paragraph.add_run().add_break()
    run = paragraph.add_run(maximum_memory)

    primary_memory_subtitle = df.iloc[237, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(primary_memory_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    primary_memory = df.iloc[238:246, 6].tolist()
    primary_memory = [hd for hd in primary_memory if pd.notna(hd)]
    paragraph.add_run().add_break()

    for hd in primary_memory:
        run = paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)

    memory_slots_subtitle = df.iloc[247, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(memory_slots_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    memory_slots = df.iloc[248:255, 6].tolist()
    memory_slots = [hd for hd in memory_slots if pd.notna(hd)]
    paragraph.add_run().add_break()

    for hd in memory_slots:
        run = paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)

    memory_footnotes = df.iloc[257:258, 6].tolist()
    memory_footnotes = [hd_footnote for hd_footnote in memory_footnotes if pd.notna(hd_footnote)]

    paragraph = doc.add_paragraph()

    for hd_footnote in memory_footnotes:
        run = paragraph.add_run(hd_footnote)

        run.font.color.rgb = RGBColor(0, 0, 255) 

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
