
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def memory_section(doc, df):

    memory_paragraph = doc.add_paragraph()
    run = memory_paragraph.add_run("MEMORY")
    run.font.size = Pt(12)
    run.bold = True
    memory_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    memory_paragraph.add_run().add_break()

    maximum_memory_subtitle = df.iloc[235, 6]
    maximum_memory_paragraph = doc.add_paragraph()
    run = maximum_memory_paragraph.add_run(maximum_memory_subtitle)
    run.bold = True
    maximum_memory_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    maximum_memory = df.iloc[236, 6]
    maximum_memory_paragraph.add_run().add_break()
    run = maximum_memory_paragraph.add_run(maximum_memory)

    primary_memory_subtitle = df.iloc[237, 6]
    primary_memory_paragraph = doc.add_paragraph()
    run = primary_memory_paragraph.add_run(primary_memory_subtitle)
    run.bold = True
    primary_memory_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    primary_memory = df.iloc[238:246, 6].tolist()
    primary_memory = [hd for hd in primary_memory if pd.notna(hd)]
    primary_memory_paragraph.add_run().add_break()

    for hd in primary_memory:
        run = primary_memory_paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)

    memory_slots_subtitle = df.iloc[247, 6]
    memory_slots_paragraph = doc.add_paragraph()

    run = memory_slots_paragraph.add_run(memory_slots_subtitle)
    run.bold = True
    memory_slots_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    memory_slots = df.iloc[248:255, 6].tolist()
    memory_slots = [hd for hd in memory_slots if pd.notna(hd)]
    memory_slots_paragraph.add_run().add_break()

    for hd in memory_slots:
        run = memory_slots_paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)

    memory_footnotes = df.iloc[257:258, 6].tolist()
    memory_footnotes = [hd_footnote for hd_footnote in memory_footnotes if pd.notna(hd_footnote)]

    memory_footnote_paragraph = doc.add_paragraph()

    for hd_footnote in memory_footnotes:
        run = memory_footnote_paragraph.add_run(hd_footnote)

        run.font.color.rgb = RGBColor(0, 0, 255) 

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
