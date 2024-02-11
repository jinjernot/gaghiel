from app.core.format.hr import *
from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
from app.core.format.table import table_column_widths
from docx.shared import Pt, Inches


def processFootnotes(doc, footnotes):
    # Create a new paragraph for footnotes with blue font color
    paragraph = doc.add_paragraph()
    for index, data in enumerate(footnotes):
        # Check if the footnote contains "Container Name" or "Wireless WAN"
        if "Container Name" in data or "Wireless WAN" in data:
            continue  # Skip this footnote if it contains unwanted values
        run = paragraph.add_run(data)
        run.font.color.rgb = RGBColor(0, 0, 153)
        
        # Check if it's not the last item before adding the line break
        if index < len(footnotes) - 1:
            run.add_break(WD_BREAK.LINE)
            
def insertTable(doc, df, html_file):
    footnotes = []  # To store footnotes temporarily
    for index, row in df.iterrows():
        # Check if the content in column 0 is "Table"
        if row[0] == "Table":
            # Add a table with 3 columns to the Word document
            table = doc.add_table(rows=1, cols=3)

            table_column_widths(table, (Inches(5), Inches(3), Inches(3)))
            
            # Populate columns 1 and 2 with values from the DataFrame
            for i in range(index + 1, len(df)):
                if df.iloc[i, 0] == "Table":
                    break  # Exit the loop when encountering the next "Table"
                elif df.iloc[i, 0] == "Footnotes":
                    # Process footnotes and store them temporarily
                    footnotes = []
                    for j in range(i + 1, len(df)):
                        if df.iloc[j, 0] == "Table":
                            break
                        footnotes.append(str(df.iloc[j, 0]))
                    # Exit the loop when footnotes are processed
                    break
                else:
                    # Populate column 1 and set text to bold
                    cell_1 = table.add_row().cells[1]
                    cell_1.text = str(df.iloc[i, 0])
                    for paragraph in cell_1.paragraphs:
                        for run in paragraph.runs:
                            run.font.bold = True

                    # Populate column 2 and set text to bold
                    cell_2 = table.rows[-1].cells[2]
                    cell_2.text = str(df.iloc[i, 1])

            # Populate column 0 and set text to bold
            cell_0 = table.cell(1, 0)
            cell_0.text = str(row[1])  # Assuming the value is in the same row
            for paragraph in cell_0.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                
            # Remove the first row from the table
            table.rows[0]._element.getparent().remove(table.rows[0]._element)
            # Replace "NaN" string values with an empty string
            
            # Process footnotes, if any
            if footnotes:
                processFootnotes(doc, footnotes)
                # Clear footnotes after processing
                footnotes = []

            # Add a paragraph break after the table
            doc.add_paragraph()

    html_table = '<table class="MsoNormalTable" cellSpacing="3" cellPadding="0" width="728" border="0">\n'
    
    for index, row in df.iterrows():
        # Check if the content in column 0 is "Table"
        if row[0] == "Table":
            # Get the starting row index for the next "Table"
            start_row_index = index + 1
            
            # Open a new table row for the header
            html_table += '<tr>'
            
            # Populate header cells for column 1 and 2
            html_table += '<th><b>{}</b></th>'.format(df.columns[0])
            html_table += '<th><b>{}</b></th>'
            
            # Close the header row
            html_table += '</tr>\n'
            
            # Determine the number of rows until the next "Table" is met
            end_row_index = start_row_index
            while end_row_index < len(df) and df.iloc[end_row_index, 0] != "Table":
                end_row_index += 1
                
            # Populate rows with values from the DataFrame
            for i in range(start_row_index, end_row_index):
                html_table += '<tr>'
                
                # Populate column 1 and set text to bold
                html_table += '<td><b>{}</b></td>'.format(df.iloc[i, 0])
                
                # Populate column 2 and set text to bold
                html_table += '<td><b>{}</b></td>'
                
                # Close the row
                html_table += '</tr>\n'
            
            # Insert the value next to the table into a new row and set text to bold
            html_table += '<tr><td colspan="2"><b>{}</b></td></tr>\n'.format(row[1])  # Assuming the value is in the same row

    # Close the table
    html_table += '</table>'

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(html_table)
