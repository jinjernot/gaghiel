
from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
import pandas as pd

def options_section(doc, file, html_file):
    """Options QS Only Section"""

    # Load xlsx
    df = pd.read_excel(file, sheet_name='QS-Only Options')

    # Add title: Options
    insertTitle(doc, "Options", html_file)

    start_col_idx = 0
    end_col_idx = 2
    start_row_idx = 0
    end_row_idx = 87

    data_range = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1]
    data_range = data_range.dropna(how='all')

    num_rows, num_cols = data_range.shape
    table = doc.add_table(rows=num_rows+1, cols=num_cols)  # Adding 1 for the header row

    table.alignment = WD_ALIGN_VERTICAL.CENTER

    # Adding table headers as the first row
    for col_idx in range(num_cols):
        header = df.columns[col_idx]
        cell = table.cell(0, col_idx)
        cell.text = header
        cell.paragraphs[0].runs[0].font.bold = True

    # Populating table cells with data
    for row_idx in range(1, num_rows+1):  # Start from the second row for data
        for col_idx in range(num_cols):
            value = data_range.iat[row_idx-1, col_idx]  # Adjust row index
            cell = table.cell(row_idx, col_idx)
            if not pd.isna(value):
                cell.text = str(value)

    html_table = '<table class="MsoNormalTable" cellSpacing="3" cellPadding="0" width="728" border="0">\n'
    for row_idx in range(num_rows):
        html_table += "  <tr>\n"
        for col_idx in range(num_cols):
            value = data_range.iat[row_idx, col_idx]
            html_table += f"    <td>{value}</td>\n"
        html_table += "  </tr>\n"
    html_table += "</table>"

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(html_table)

    # Insert HR
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)