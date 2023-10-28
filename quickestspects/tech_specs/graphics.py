
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
import pandas as pd

def graphics_section(doc, df):

    graphics_paragraph = doc.add_paragraph()
    run = graphics_paragraph.add_run("GRAPHICS")
    run.font.size = Pt(12)
    run.bold = True
    graphics_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    graphics_paragraph.add_run().add_break()

    integrated_paragraph = doc.add_paragraph()
    run = integrated_paragraph.add_run("Integrated")
    run.bold = True
    integrated_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    integrated_paragraph.add_run().add_break()

    integrated = df.iloc[103:108, 6].tolist()
    integrated = [gfx for gfx in integrated if pd.notna(gfx)]
    integrated_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for gfx in integrated:
        run = integrated_paragraph.add_run(gfx)

    discrete_paragraph = doc.add_paragraph()
    run = discrete_paragraph.add_run("Discrete")
    run.bold = True
    discrete_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    discrete_paragraph.add_run().add_break()

    discrete = df.iloc[110:111, 6].tolist()
    discrete = [gfx for gfx in discrete if pd.notna(gfx)]
    discrete_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for gfx in discrete:
        run = discrete_paragraph.add_run(gfx)

    supports_paragraph = doc.add_paragraph()
    run = supports_paragraph.add_run("Supports")
    run.bold = True
    supports_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    supports_paragraph.add_run().add_break()

    supports = df.iloc[112:116, 6].tolist()
    supports = [gfx for gfx in supports if pd.notna(gfx)]
    supports_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for gfx in supports:
        run = supports_paragraph.add_run(gfx)

    graphics_footnotes = df.iloc[117:121, 6].tolist()
    graphics_footnotes = [gfx_footnote for gfx_footnote in graphics_footnotes if pd.notna(gfx_footnote)]

    # Create a new paragraph
    graphics_footnote_paragraph = doc.add_paragraph()

    # Add the data from the list to the paragraph
    for gfx_footnote in graphics_footnotes:
        run = graphics_footnote_paragraph.add_run(gfx_footnote)

        # Set the font color to blue
        run.font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
