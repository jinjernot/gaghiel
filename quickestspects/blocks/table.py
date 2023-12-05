from quickestspects.format.hr import *

def insertTable(doc, df, txt_file):
    for index, row in df.iterrows():
        # Check if the content in column 0 is "Table"
        if row[0] == "Table":
            # Get the starting row index for the next "Table"
            start_row_index = index + 1
            
            # Add a table with 3 columns to the Word document
            table = doc.add_table(rows=1, cols=3)
            
            # Determine the number of rows until the next "Table" is met
            end_row_index = start_row_index
            while end_row_index < len(df) and df.iloc[end_row_index, 0] != "Table":
                end_row_index += 1
            
            # Populate columns 1 and 2 with values from the DataFrame
            for i in range(start_row_index, end_row_index):
                # Populate column 1 and set text to bold
                cell_1 = table.add_row().cells[1]
                cell_1.text = str(df.iloc[i, 0])
                for paragraph in cell_1.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

                # Populate column 2 and set text to bold
                cell_2 = table.rows[-1].cells[2]
                cell_2.text = str(df.iloc[i, 1])
                        
            # Insert the value next to the table into the second row of column 0 and set text to bold
            cell_0 = table.cell(1, 0)
            cell_0.text = str(row[1])  # Assuming the value is in the same row
            for paragraph in cell_0.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                # Insert HR
            insertHR(doc.add_paragraph(), thickness=3)

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
            html_table += '<th><b>{}</b></th>'.format(df.columns[1])
            
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
                html_table += '<td><b>{}</b></td>'.format(df.iloc[i, 1])
                
                # Close the row
                html_table += '</tr>\n'
            
            # Insert the value next to the table into a new row and set text to bold
            html_table += '<tr><td colspan="2"><b>{}</b></td></tr>\n'.format(row[1])  # Assuming the value is in the same row

    # Close the table
    html_table += '</table>'

    with open(txt_file, 'a', encoding='utf-8') as txt:
        txt.write(html_table)
