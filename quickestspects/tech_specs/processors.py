
from quickestspects.format.hr import insertHR
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,  WD_ALIGN_VERTICAL
from docx.shared import RGBColor
from docx.enum.text import WD_BREAK
import pandas as pd

def processors_section(doc, df):

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("PROCESSORS")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

 # Define the column indices for the range G54:M60
    start_col_idx = 6  # Column G
    end_col_idx = 12  # Column M
    start_row_idx = 52
    end_row_idx = 60

    # Select the data range using column indices
    data_range = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1]

    # Remove rows with all NaN values
    data_range = data_range.dropna(how='all')

    # Create a table in the document with the same number of rows and columns as the data
    num_rows, num_cols = data_range.shape
    table = doc.add_table(rows=num_rows, cols=num_cols)

    # Set table alignment
    table.alignment = WD_ALIGN_VERTICAL.CENTER

    # Populate the table with data and handle NaN values
    for row_idx in range(num_rows):
        for col_idx in range(num_cols):
            value = data_range.iat[row_idx, col_idx]
            cell = table.cell(row_idx, col_idx)

            if not pd.isna(value):
                cell.text = str(value)

    # Make the first row bold
    for cell in table.rows[0].cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
    

    run.add_break(WD_BREAK.LINE)

    processors_footnotes = df.iloc[73:80, 6].tolist()
    processors_footnotes = [os for os in processors_footnotes if pd.notna(os)]
    
    # Create a new paragraph
    processor_footnote_paragraph = doc.add_paragraph()

    # Add the data from the list to the paragraph
    for pro_footnote in processors_footnotes:
        run = processor_footnote_paragraph.add_run(pro_footnote)
        run.add_break(WD_BREAK.LINE)
        # Set the font color to blue
        run.font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue

    run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)