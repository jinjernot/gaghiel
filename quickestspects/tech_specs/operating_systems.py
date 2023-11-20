
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, RGBColor
import pandas as pd

def operating_systems_section(doc, txt_file, df):

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("OPERATING SYSTEMS")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    with open(txt_file, 'a') as txt:
        txt.write("<h1>OPERATING SYSTEMS</h1>\n")

    operating_systems = df.iloc[19:30, 6].tolist()

    operating_systems = [os for os in operating_systems if pd.notna(os)]

    total_rows = (len(operating_systems))

    os_table = doc.add_table(rows=total_rows, cols=2)
    os_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    col_index = 1 
    for row_index in range(total_rows):
        list_index = row_index
        if list_index < len(operating_systems):
            os_table.cell(row_index, col_index).text = str(operating_systems[list_index])
    
    preinstalled_text = df.iloc[12, 6]
    preinstalled = os_table.cell(0, 0)
    preinstalled.text = preinstalled_text
    run = preinstalled.paragraphs[0].runs[0]
    run.bold = True

    html_table = '<table class="MsoNormalTable" cellSpacing="3" cellPadding="0" width="728" border="0">\n'

    html_table += f'<tr>\n<td><strong>{preinstalled_text}</strong></td>\n</tr>\n'

    for os in operating_systems:
        html_table += f'<tr>\n<td></td>\n<td>{os}</td>\n</tr>\n'
    # Closing HTML tags
    html_table += '</table>\n'
    with open(txt_file, 'a') as txt:
            txt.write(html_table)



    operating_systems_footnotes = df.iloc[31:36, 6].tolist()
    operating_systems_footnotes = [os for os in operating_systems_footnotes if pd.notna(os)]

    footnote_paragraph = doc.add_paragraph()

    for os_footnote in operating_systems_footnotes:
        run = footnote_paragraph.add_run(os_footnote)

        run.font.color.rgb = RGBColor(0, 0, 255) 

        run.add_break(WD_BREAK.LINE)


    html_footnotes = '<div style="color: blue;">\n'

    for os_footnote in operating_systems_footnotes:
        html_footnotes += f'  <span>{os_footnote}</span>\n'
        
    html_footnotes += '</div>\n'

    with open(txt_file, 'a') as txt:
            txt.write(html_footnotes)

    insertHR(doc.add_paragraph(), thickness=3)

    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)