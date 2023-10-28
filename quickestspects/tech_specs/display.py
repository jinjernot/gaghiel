
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
import pandas as pd

def display_section(doc, df):

    display_paragraph = doc.add_paragraph()
    run = display_paragraph.add_run("DISPLAY")
    run.font.size = Pt(12)
    run.bold = True
    display_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    display_paragraph.add_run().add_break()

    non_touch_paragraph = doc.add_paragraph()
    run = non_touch_paragraph.add_run("Non-Touch")
    run.bold = True
    non_touch_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    non_touch_paragraph.add_run().add_break()

    non_touch = df.iloc[131:142, 6].tolist()
    non_touch = [disp for disp in non_touch if pd.notna(disp)]
    non_touch_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in non_touch:
        run = non_touch_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    touch_paragraph = doc.add_paragraph()
    run = touch_paragraph.add_run("Touch")
    run.bold = True
    touch_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    touch_paragraph.add_run().add_break()

    touch = df.iloc[146:149, 6].tolist()
    touch = [disp for disp in touch if pd.notna(disp)]
    touch_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in touch:
        run = touch_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    dp_paragraph = doc.add_paragraph()
    run = dp_paragraph.add_run("DisplayPort™ 1.2")
    run.bold = True
    dp_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    dp_paragraph.add_run().add_break()

    displayport = df.iloc[156:157, 6].tolist()
    displayport = [disp for disp in displayport if pd.notna(disp)]
    dp_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in displayport:
        run = dp_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    display_footnotes = df.iloc[186:191, 6].tolist()
    display_footnotes = [disp_footnote for disp_footnote in display_footnotes if pd.notna(disp_footnote)]

    # Create a new paragraph
    graphics_footnote_paragraph = doc.add_paragraph()

    # Add the data from the list to the paragraph
    for disp_footnote in display_footnotes:
        run = graphics_footnote_paragraph.add_run(disp_footnote)

        # Set the font color to blue
        run.font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
