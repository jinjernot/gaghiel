import os
from docx.shared import Pt, Inches
from docx.enum.text import WD_BREAK
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
import pandas as pd


def callout_section(doc, df, prod_name, imgs_path):

    prodname_paragraph = doc.add_paragraph()
    run = prodname_paragraph.add_run(prod_name)
    run.font.size = Pt(12)
    run.bold = True
    prodname_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    img_path = os.path.join(imgs_path, 'placeholder-image.png')
    img_path2 = os.path.join(imgs_path, 'placeholder-image.png')

    doc.add_picture(img_path, width=Inches(6))
    
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Front")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Define the 'calloutfront_' tags you want to process
    callouts = df.iloc[11:31, 6].tolist()

    # Filter out NaN values
    callouts = [tag for tag in callouts if pd.notna(tag)]

    # Create a new list to store the previous column data
    previous_column_data = df.iloc[11:31, 5].tolist()

    # Calculate the number of rows needed
    total_rows = (len(callouts) + 1) // 2  # Adding 1 to round up if there's an odd number of tags

    # Create the table with the dynamically determined number of rows and 4 columns
    callout_table = doc.add_table(rows=total_rows, cols=2)
    callout_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Populate the table
    for row_index in range(total_rows):
        for col_index in range(2):
            list_index = row_index * 2 + col_index
            if list_index < len(callouts):
                callout_table.cell(row_index, col_index).text = str(callouts[list_index])

    # Center the table cells
        for row in callout_table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in cell.paragraphs[0].runs:
                    run.font.size = Pt(10) 

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    doc.add_picture(img_path2, width=Inches(6))

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Back")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Define the 'calloutback_' tags you want to process
    tags_to_process_back = df.iloc[40:60, 6].tolist()

    # Filter out NaN values
    filtered_tags2 = [tag for tag in tags_to_process_back if pd.notna(tag)]

    # Calculate the number of rows needed
    total_rows = (len(filtered_tags2) + 1) // 2  # Adding 1 to round up if there's an odd number of tags

    # Create the table with the dynamically determined number of rows
    callout_table2 = doc.add_table(rows=total_rows, cols=2)
    callout_table2.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Populate the table
    for row_index in range(total_rows):
        for col_index in range(2):
            list_index = row_index * 2 + col_index
            if list_index < len(filtered_tags2):
                callout_table2.cell(row_index, col_index).text = str(filtered_tags2[list_index])

    # Center the table cells
    for row in callout_table2.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(10) 

    doc.add_page_break() 
    section = doc.sections[-1]
    section.start_type
    section.start_type = WD_SECTION.CONTINUOUS