
from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *
from docx.shared import Inches
from docx.enum.text import WD_BREAK
from app.core.format.table import table_column_widths

def docking_section(doc, html_file, df):
    """Docking Table"""

    # Add title: DOCKING
    insertTitle(doc, "Docking (Sold Separately)", html_file)

    for index, row in df.iterrows():
        # Check if the content in column 0 is "Table"
        if row[1] == "Docking":
            # Add a table with 2 columns to the Word document
            table = doc.add_table(rows=1, cols=2)
            table_column_widths(table, (Inches(3), Inches(5)))
            
            # Populate columns 0 and 1 with values from the DataFrame
            for i in range(index + 1, len(df)):
                if df.iloc[i, 0] == "Container Name":
                    break  # Exit the loop when encountering the next "Table"
                else:
                    # Populate column 0 and set text to bold
                    cell_0 = table.add_row().cells[0]
                    cell_0.text = str(df.iloc[i, 0])
                    for paragraph in cell_0.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True

                    # Populate column 1 and set text to bold
                    cell_1 = table.rows[-1].cells[1]
                    cell_1.text = str(df.iloc[i, 1])

            # Remove the first row from the table
            table.rows[0]._element.getparent().remove(table.rows[0]._element)
            
            # Add a paragraph break after the table
            doc.add_paragraph()
    # Insert HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(html_file)
    
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
