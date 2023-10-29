
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def system_unit_section(doc, xlsx_file, df):

    df = pd.read_excel(xlsx_file, sheet_name='QS-Only System Unit')

    audio_paragraph = doc.add_paragraph()
    run = audio_paragraph.add_run("SYSTEM UNIT")
    run.font.size = Pt(12)
    run.bold = True
    audio_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    audio_paragraph.add_run().add_break()

     # Define the column indices for the range G54:M60
    start_col_idx = 5  # Column G
    end_col_idx = 6 # Column M
    start_row_idx = 10
    end_row_idx = 52

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


    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
