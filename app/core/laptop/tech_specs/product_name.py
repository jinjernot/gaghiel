from app.core.format.hr import *

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

import pandas as pd

def product_name_section(doc, file, html_file):
    """Product name section"""

    # Get product name
    df = pd.read_excel(file, sheet_name='Callouts')
    #df = pd.read_excel(file.stream, sheet_name='Tech Specs & QS Features', engine='openpyxl')
    prod_name = df.columns[1]


    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Product Name")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    html_content = f'''
<p class="MsoNormal" style="LINE-HEIGHT: 115%"><span lang="EN-US" style="COLOR: #000099">&nbsp;</span></p></div></div><qsnav heading="Technical Specifications"><a name="Technical Specifications"></a>\n
<p class="title">Technical Specifications</p></qsnav>
<div class="section">
<h2 style="MARGIN-TOP: 8pt; LINE-HEIGHT: 115%"><span lang="EN-US">PRODUCT NAME</span></h2>
<table class="MsoNormalTable" cellSpacing="3" cellPadding="0" width="720" border="0">
<tbody>
<tr>
<td style="WIDTH: 537.05pt; PADDING-BOTTOM: 0.75pt; PADDING-TOP: 0.75pt; PADDING-LEFT: 0.75pt; PADDING-RIGHT: 0.75pt" vAlign="top" width="716">
<h2 style="MARGIN-TOP: 0cm; LINE-HEIGHT: 115%"><span lang="EN-US" style="FONT-SIZE: 10pt; FONT-VARIANT: normal !important; FONT-WEIGHT: normal; COLOR: black; LETTER-SPACING: 0pt; LINE-HEIGHT: 115%">{prod_name}</span></h2></td></tr>
'''

    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(html_content)


    paragraph = doc.add_paragraph(prod_name)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(html_file)