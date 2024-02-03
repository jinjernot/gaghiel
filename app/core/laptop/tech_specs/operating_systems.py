from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK
import pandas as pd

def operating_systems_section(doc, html_file, df):
    """Operating system techspecs section"""

    # Add the title: OPERATING SYSTEMS
    insertTitle(doc, "OPERATING SYSTEMS", html_file)
    operating_systems = df.iloc[3:8, 1].tolist()
    operating_systems = [os for os in operating_systems if pd.notna(os)]

    total_rows = (len(operating_systems))
    
    total_document_width = doc.sections[0].page_width 
    first_column_width = int(0.25 * total_document_width)
    second_column_width = total_document_width - first_column_width

    os_table = doc.add_table(rows=total_rows, cols=2)
    os_table.columns[0].width = first_column_width
    os_table.columns[1].width = second_column_width
    #os_table.alignment = WD_TABLE_ALIGNMENT.LEFT

    col_index = 1 
    for row_index in range(total_rows):
        list_index = row_index
        if list_index < len(operating_systems):
            os_table.cell(row_index, col_index).text = str(operating_systems[list_index])
    
    preinstalled_text = df.iloc[2, 1]
    preinstalled = os_table.cell(0, 0)
    preinstalled.text = preinstalled_text
    run = preinstalled.paragraphs[0].runs[0]
    run.bold = True

    html_table = '<table class="MsoNormalTable" cellSpacing="3" cellPadding="0" width="720" border="0">\n'

    html_table += f'<tbody><tr>\n<td style="WIDTH: 102.2pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="136"><p class="MsoNormal" style="LINE-HEIGHT: 115%"><b><span lang="EN-US">{preinstalled_text}</span></b></p></td>\n</tr>\n'

    for os in operating_systems:
        html_table += f'<tr>\n<td></td>\n<td style="WIDTH: 416.1pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="555"><p class="MsoNormal" style="LINE-HEIGHT: 115%"><b><span lang="EN-US">{os}</span></p></td>\n</tr>\n'
    html_table += '</table>\n'
    with open(html_file, 'a', encoding='utf-8') as txt:
            txt.write(html_table)
            txt.write('<p class="MsoNormal" style="LINE-HEIGHT: 115%"></p></td></tr></tbody></table>\n')
            txt.write('<tr style="HEIGHT: 13.6pt">\n')
            txt.write('<td style="HEIGHT: 13.6pt; WIDTH: 537.05pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" width="716" colSpan="3">\n')
            
    # Footnotes
    insertFootnote(doc, html_file, df, slice(10, 11), 1)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(html_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)