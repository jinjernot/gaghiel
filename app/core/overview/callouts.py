from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.shared import Pt, Inches
import pandas as pd
import os


def callout_section(doc, html_file, prod_name, imgs_path, df):
    """Add Callout Section"""

    # Add the product name
    prodname_paragraph = doc.add_paragraph()
    run = prodname_paragraph.add_run(prod_name)
    run.font.size = Pt(12)
    run.bold = True
    prodname_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # add HTML Headers
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(f"<html><head><title>{prod_name}</title>\n")
        txt.write('<meta content="text/html; charset=utf-8" http-equiv="Content-Type">\n')
        txt.write('<meta name="Generator" content="Microsoft Word 15 (filtered)">\n')
        txt.write("</head>\n")

        #Add title and product name
        txt.write("<h1><b>Overview</h1></b>\n")
        txt.write(f"<p style='font-size:14pt;'><strong>{prod_name}</strong></p>\n")

    # Image paths
    img_path = os.path.join(imgs_path, 'image001.png')
    img_path2 = os.path.join(imgs_path, 'image002.png')

    # Image HTML Tags
    img_html_code = f'<img src="{img_path}" alt="Product Image" width="702" height="561">'
    img_html_code2 = f'<img src="{img_path2}" alt="Product Image" width="702" height="561">'

    # Left image HTML subtitle
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(img_html_code + '\n')
        txt.write("<b><p>Front</p></b>\n")
    # Add  left image to docx
    doc.add_picture(img_path, width=Inches(6))

    # Add Left subtitle
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Front")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    start_col_idx = 1
    end_col_idx = 4
    start_row_idx = 4
    end_row_idx = 9

    data_range = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1]
    data_range = data_range.dropna(how='all')

    num_rows, num_cols = data_range.shape
    table = doc.add_table(rows=num_rows, cols=num_cols)

    table.alignment = WD_ALIGN_VERTICAL.CENTER

    for row_idx in range(num_rows):
        for col_idx in range(num_cols):
            value = data_range.iat[row_idx, col_idx]
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

    # add HTML <hr>
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # add docx image
    doc.add_picture(img_path2, width=Inches(6))

    # Right image HTML subtitle
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(img_html_code2 + '\n')
        txt.write("<b><p>Right</p></b>\n")

    # Add Right subtitle
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Right")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    start_col_idx = 1
    end_col_idx = 4
    start_row_idx = 10
    end_row_idx = 22

    data_range = df.iloc[start_row_idx:end_row_idx+1, start_col_idx:end_col_idx+1]
    data_range = data_range.dropna(how='all')

    num_rows, num_cols = data_range.shape
    table = doc.add_table(rows=num_rows, cols=num_cols)

    table.alignment = WD_ALIGN_VERTICAL.CENTER

    for row_idx in range(num_rows):
        for col_idx in range(num_cols):
            value = data_range.iat[row_idx, col_idx]
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
    # add HTML <hr>
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    doc.add_page_break()
    section = doc.sections[-1]
    section.start_type
    section.start_type = WD_SECTION.CONTINUOUS
