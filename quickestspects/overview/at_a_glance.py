from docx.enum.text import WD_BREAK,WD_ALIGN_PARAGRAPH
import pandas as pd

def ataglance_section(doc, df):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("At a Glance")
    run.bold = True

    at_a_glance = df.iloc[70:85, 6].tolist()
    
    for item in at_a_glance:
        if pd.notna(item):  # Check for NaN values
            # Add each item from at_a_glance as a bulleted list item
            doc.add_paragraph(item, style='List Bullet')
    
     # Add the note
    note_paragraph = doc.add_paragraph("NOTE: See important legal disclosures for all listed specs in their respective features sections")
    note_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for run in note_paragraph.runs:
        run.bold = True

    # Add a page break after the bulleted list
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)