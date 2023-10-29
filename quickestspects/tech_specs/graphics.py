
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def graphics_section(doc, df):

    graphics_paragraph = doc.add_paragraph()
    run = graphics_paragraph.add_run("GRAPHICS")
    run.font.size = Pt(12)
    run.bold = True
    graphics_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    graphics_paragraph.add_run().add_break()


    integrated_subtitle = df.iloc[102, 6]
    integrated_paragraph = doc.add_paragraph()
    run = integrated_paragraph.add_run(integrated_subtitle)
    run.bold = True
    integrated_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    integrated = df.iloc[103:108, 6].tolist()
    integrated = [gfx for gfx in integrated if pd.notna(gfx)]
    integrated_paragraph.add_run().add_break()

    for gfx in integrated:
        run = integrated_paragraph.add_run(gfx)


    discrete_subtitle = df.iloc[108, 6]
    discrete_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = discrete_paragraph.add_run(discrete_subtitle)
    run.bold = True
    discrete_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    integrated_paragraph.add_run().add_break()

    discrete = df.iloc[110:111, 6].tolist()
    discrete = [gfx for gfx in discrete if pd.notna(gfx)]
    discrete_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for gfx in discrete:
        run = discrete_paragraph.add_run(gfx)

    # Assuming that df.iloc[102, 6] contains the text you want to use
    supports_subtitle = df.iloc[111, 6]
    
    # Create a new paragraph in your Word document
    supports_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = supports_paragraph.add_run(supports_subtitle)
    run.bold = True
    supports_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    supports = df.iloc[112:116, 6].tolist()
    supports = [gfx for gfx in supports if pd.notna(gfx)]
    supports_paragraph.add_run().add_break()

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
