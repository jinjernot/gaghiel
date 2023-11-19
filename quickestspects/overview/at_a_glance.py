from docx.enum.text import WD_BREAK,WD_ALIGN_PARAGRAPH
import pandas as pd

def ataglance_section(doc, txt_file, df):

    # Add a title to the Word document
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("At a Glance")
    run.bold = True

    with open(txt_file, 'a') as txt:
        txt.write("<h1>At a Glance</h1>\n")

    # Extract data from the DataFrame
    at_a_glance = df.iloc[70:85, 6].tolist()
    
    for item in at_a_glance:
        if pd.notna(item):  # Check for NaN values
            # Add each item from at_a_glance as a bulleted list item
            doc.add_paragraph(item, style='List Bullet')


            # Append the item to the text file
            with open(txt_file, 'a') as txt:
                txt.write(f"<p>{item}</p>\n")

     # Add the note
    note_paragraph = doc.add_paragraph("NOTE: See important legal disclosures for all listed specs in their respective features sections")
    note_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Append the note to the text file
    with open(txt_file, 'a') as txt:
        txt.write("<b>NOTE: See important legal disclosures for all listed specs in their respective features sections</b>\n")
        
    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    for run in note_paragraph.runs:
        run.bold = True

    # Add a page break after the bulleted list
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)