from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
import pandas as pd

def insertParagraph(doc, html_file, df, iloc_row, iloc_column):
    data = df.iloc[iloc_row, iloc_column]
    paragraph = doc.add_paragraph()
    paragraph.add_run(data)

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(f"<p>{data}</p>\n")

def processFootnotes(doc, footnotes):
    # Create a new paragraph for footnotes with blue font color
    paragraph = doc.add_paragraph()
    for index, data in enumerate(footnotes):
        run = paragraph.add_run(data)
        run.font.color.rgb = RGBColor(0, 0, 153)
        
        # Check if it's not the last item before adding the line break
        if index < len(footnotes) - 1:
            run.add_break(WD_BREAK.LINE)

def insertList(doc, html_file, df, start_value):
    # Check if start_value exists in the DataFrame
    if start_value not in df.iloc[:, 1].tolist():
        print(f"Error: '{start_value}' not found in DataFrame.")
        return
    
    # Find the index of the start_value
    start_index = df.index[df.iloc[:, 1] == start_value].tolist()[0]
    print("Start index:", start_index)

    # Find the index of the next occurrence of "Value" after the start_value index
    next_value_indices = df.iloc[start_index:, 1][df.iloc[start_index:, 1] == 'Value'].index.tolist()
    if not next_value_indices:
        print("Error: 'Value' not found after", start_value)
        return
    next_value_index = next_value_indices[0]
    print("Next value index:", next_value_index)
    
    # Get the data between the start_value and "Value" from the second column
    items = df.iloc[start_index:next_value_index, 1].tolist()
    print("Items:", items)
    
    # Remove "Footnotes" if it exists and separate values after it
    if 'Footnotes' in items:
        footnotes_index = items.index('Footnotes')
        items = items[:footnotes_index]
        footnotes = df.iloc[footnotes_index + start_index + 1:next_value_index, 1].tolist()
    else:
        footnotes = []
       
    # Create a paragraph for items before "Footnotes"
    paragraph = doc.add_paragraph()
    
    # Add each item to the paragraph with a line break
    for index, data in enumerate(items):
        run = paragraph.add_run(data)
        
        # Check if it's not the last item before adding the line break
        if index < len(items) - 1:
            run.add_break(WD_BREAK.LINE)
    
    # Process footnotes after the items have been added
    if footnotes:
        processFootnotes(doc, footnotes)

def insertFootnote(doc, html_file, df, iloc_range, iloc_column):
    # Get the data
    footnote = df.iloc[iloc_range, iloc_column].tolist()

    # Create a paragraph
    paragraph = doc.add_paragraph()

    # Add each footnote to the paragraph with a line break, set font
    for index, note in enumerate(footnote):
        run = paragraph.add_run(note)
        # Set color to Blue
        run.font.color.rgb = RGBColor(0, 0, 153)
        
        # Check if it's not the last item before adding the line break
        if index < len(footnote) - 1:
            run.add_break(WD_BREAK.LINE)

    html_footnotes = '<tr>\n'
    for index, note in enumerate(footnote):
        html_footnotes += f'<p class="MsoNormal" style="LINE-HEIGHT: 115%"><span lang="EN-US" style="COLOR: #000099">{note}</span></p>'
        # Check if it's not the last item before closing the <span> tag
        if index < len(footnote) - 1:
            html_footnotes += '\n'
    html_footnotes += '</td></tr>\n'

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write('<tr>')
        txt.write('<td style="WIDTH: 15.75pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="21">')
        txt.write('<p class="MsoNormal" style="LINE-HEIGHT: 115%"><span lang="EN-US" style="COLOR: #000099">&nbsp;</span></p></td>')
        txt.write('<td style="WIDTH: 519.8pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="693" colSpan="2">')
        txt.write('<p class="MsoNormal" style="LINE-HEIGHT: 115%"><span lang="EN-US" style="COLOR: black">&nbsp;</span></p></td></tr>')
        txt.write('<tr>')
        txt.write('<td style="WIDTH: 15.75pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="21">')
        txt.write('<td style="WIDTH: 519.8pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="693" colSpan="2">')
        txt.write(html_footnotes)
