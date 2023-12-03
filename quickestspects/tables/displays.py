
from quickestspects.blocks.paragraph import *
from quickestspects.blocks.title import *
from quickestspects.format.hr import *

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def displays_section(doc, xlsx_file, txt_file):
    """Displaystable"""

    # Load xlsx
    df = pd.read_excel(xlsx_file, sheet_name='QS-Only Displays')

    # Add tible: Displays
    insertTitle(doc, "Displays", txt_file)

    for index, row in df.iterrows():
        # Check if the content in column 0 is "Table"
        if row[0] == "Table":
            # Get the value next to the table (assuming it's in the next column)
            value_next_to_table = df.iloc[index, 1]
            
            # Add a table with 3 rows to the Word document
            table = doc.add_table(rows=17, cols=3)
            
            # Populate the table
            # Insert the value next to the table into the first cell of column 0
            table.cell(0, 0).text = str(value_next_to_table)




    # Insert HR
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
