from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt, RGBColor
import pandas as pd

def display_section(doc, df):

    display_paragraph = doc.add_paragraph()
    run = display_paragraph.add_run("DISPLAY")
    run.font.size = Pt(12)
    run.bold = True
    display_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    display_paragraph.add_run().add_break()

    # Get the Subtitle
    non_touch_subtitle = df.iloc[130, 6]
    
    # Create a new paragraph in your Word document
    non_touch_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = non_touch_paragraph.add_run(non_touch_subtitle)
    run.bold = True
    non_touch_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    non_touch_paragraph.add_run().add_break()

    non_touch = df.iloc[131:142, 6].tolist()
    non_touch = [disp for disp in non_touch if pd.notna(disp)]
    non_touch_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in non_touch:
        run = non_touch_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    # Get the Subtitle
    touch_subtitle = df.iloc[145, 6]
    
    # Create a new paragraph in your Word document
    touch_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = touch_paragraph.add_run(touch_subtitle)
    run.bold = True
    touch_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    touch_paragraph.add_run().add_break()

    touch = df.iloc[146:149, 6].tolist()
    touch = [disp for disp in touch if pd.notna(disp)]
    touch_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in touch:
        run = touch_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    # Get the Subtitle
    displayport_subtitle = df.iloc[155, 6]
    
    # Create a new paragraph in your Word document
    displayport_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = displayport_paragraph.add_run(displayport_subtitle)
    run.bold = True
    displayport_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    displayport_paragraph.add_run().add_break()

    displayport = df.iloc[156:157, 6].tolist()
    displayport = [disp for disp in displayport if pd.notna(disp)]
    dp_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in displayport:
        run = dp_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

        # Get the Subtitle
    display_support_subtitle = df.iloc[158, 6]
    
    # Create a new paragraph in your Word document
    display_support_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = display_support_paragraph.add_run(display_support_subtitle)
    run.bold = True
    display_support_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    display_support_paragraph.add_run().add_break()

    display_support = df.iloc[159:161, 6].tolist()
    display_support = [disp for disp in display_support if pd.notna(disp)]
    display_support_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in display_support:
        run = display_support_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

    # Get the Subtitle
    display_size_subtitle = df.iloc[161, 6]
    
    # Create a new paragraph in your Word document
    display_size_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = display_size_paragraph.add_run(display_size_subtitle)
    run.bold = True
    display_size_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Add a line break
    display_size_paragraph.add_run().add_break()

    display_size = df.iloc[162:164, 6].tolist()
    display_size = [disp for disp in display_size if pd.notna(disp)]
    display_size_paragraph = doc.add_paragraph()

     # Add the data from the list to the paragraph
    for disp in display_size:
        run = display_size_paragraph.add_run(disp)
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
