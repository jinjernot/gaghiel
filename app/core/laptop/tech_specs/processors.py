from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK

import pandas as pd

def processors_section(doc, file, html_file):
    """Processors techspecs section"""

    df = pd.read_excel(file, sheet_name='QS-Only Processors')
    #df = pd.read_excel(file.stream, sheet_name='QS-Only Processors', engine='openpyxl')

    # Add the title: PROCESSORS
    insert_title(doc, "Processors", html_file)

    # Define the criteria to filter the rows
    criteria_values = ["Processor [3,4]", "Cores", "Threads", "L3 Cache", "Max Boost Frequency [5]", "Base Frequency", "Pro"]

    # Filter the dataframe based on the values in the third row
    third_row = df.iloc[1]  # Selecting the third row
    filtered_df = df.loc[:, third_row.isin(criteria_values)]  # Filtering columns based on criteria

    # Remove the first row
    filtered_df = filtered_df.iloc[1:]

    # Replace "NaN" string values with an empty string
    filtered_df = filtered_df.fillna('')

    # Convert filtered dataframe to a list of lists (data) for table
    data = filtered_df.values.tolist()

    # Add the data as a table to the document
    table = doc.add_table(rows=1, cols=len(data[0]))
    table.autofit = True

    # Add table data
    for row in data:
        row_cells = table.add_row().cells
        for i, cell in enumerate(row):
            row_cells[i].text = str(cell)

    # Remove the first row
    if len(table.rows) > 1:  # Ensure there are rows to delete
        table.rows[0]._element.getparent().remove(table.rows[0]._element)

    # HR
    insert_horizontal_line(doc.add_paragraph(), thickness=3)
    insert_html_horizontal_line(html_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)