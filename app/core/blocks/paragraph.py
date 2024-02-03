from docx.enum.text import WD_BREAK
from docx.shared import RGBColor
import pandas as pd

def insertParagraph(doc, html_file, df, iloc_row, iloc_column):
    data = df.iloc[iloc_row, iloc_column]
    paragraph = doc.add_paragraph()
    paragraph.add_run(data)

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(f"<p>{data}</p>\n")

def insertList(doc, html_file, df, iloc_range, iloc_column):
    # Get the data
    items = df.iloc[iloc_range, iloc_column].tolist()
    # Remove N/A 
    items = [data for data in items if pd.notna(data)]

    # Create a paragraph
    paragraph = doc.add_paragraph()

    # Add each item to the paragraph with a line break
    for index, data in enumerate(items):
        run = paragraph.add_run(data)

        # Check if it's not the last item before adding the line break
        if index < len(items) - 1:
            run.add_break(WD_BREAK.LINE)

        with open(html_file, 'a', encoding='utf-8') as txt:
            txt.write(f"<p>{data}</p>")
            # Check if it's not the last item before adding the line break in HTML
            if index < len(items) - 1:
                txt.write('\n')

def insertFootnote(doc, html_file, df, iloc_range, iloc_column):
    # Get the data
    footnote = df.iloc[iloc_range, iloc_column].tolist()
    #Remove N/A 
    footnote  = [note for note in footnote if pd.notna(note)]

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
    html_footnotes += '</tr>\n'

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write('<td style="WIDTH: 15.75pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="21">')
        txt.write('<p class="MsoNormal" style="LINE-HEIGHT: 115%"><span lang="EN-US" style="COLOR: #000099">&nbsp;</span></p></td>')
        txt.write('<td style="WIDTH: 519.8pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="693" colSpan="2">')

        txt.write(html_footnotes)
