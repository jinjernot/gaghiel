
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def storage_section(doc, df):

    storage_paragraph = doc.add_paragraph()
    run = storage_paragraph.add_run("STORAGE AND DRIVES")
    run.font.size = Pt(12)
    run.bold = True
    storage_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    storage_paragraph.add_run().add_break()


    primary_storage_subtitle = df.iloc[200, 6]
    primary_storage_paragraph = doc.add_paragraph()
    run = primary_storage_paragraph.add_run(primary_storage_subtitle)
    run.bold = True
    primary_storage_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    primary_storage = df.iloc[201:207, 6].tolist()
    primary_storage = [hd for hd in primary_storage if pd.notna(hd)]
    primary_storage_paragraph.add_run().add_break()

    for hd in primary_storage:
        run = primary_storage_paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)


    m2_storage_subtitle = df.iloc[207, 6]
    m2_storage_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = m2_storage_paragraph.add_run(m2_storage_subtitle)
    run.bold = True
    m2_storage_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    m2_storage = df.iloc[208:221, 6].tolist()
    m2_storage = [hd for hd in m2_storage if pd.notna(hd)]
    m2_storage_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for hd in m2_storage:
        run = m2_storage_paragraph.add_run(hd)
        run.add_break(WD_BREAK.LINE)

    storage_footnotes = df.iloc[222:226, 6].tolist()
    storage_footnotes = [hd_footnote for hd_footnote in storage_footnotes if pd.notna(hd_footnote)]

    # Create a new paragraph
    storage_footnote_paragraph = doc.add_paragraph()

    # Add the data from the list to the paragraph
    for hd_footnote in storage_footnotes:
        run = storage_footnote_paragraph.add_run(hd_footnote)

        # Set the font color to blue
        run.font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
