import pandas as pd
import glob
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os

def loadxlsx():
    """Load the xlsx file"""
    folder_path = "./xlsx/"
    imgs_path = "./imgs/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")

    for xlsx_file in xlsx_files: #loop through all the files
        createdocx(xlsx_file, imgs_path)

def createdocx(xlsx_file, imgs_path):
    """Create the Quickestspecs"""

    df = pd.read_excel(xlsx_file)
    df = df[df['4RA85F [Product]'] != '##BLANK##']

    img_path = os.path.join(imgs_path, 'c08518669.png')
    prod_name = df.loc[df['Tag'] == 'prodname', '4RA85F [Product]'].iloc[0]

    doc = Document()

    header = doc.sections[0].header
    header_table = header.add_table(rows=1, cols=2, width=Inches(8)) 
    header_table.rows[0].height = Inches(.5)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.allow_autofit = False

    cell = header_table.cell(0, 0)
    cell_paragraph = cell.add_paragraph()
    cell_paragraph.add_run("Quickestspecs").font.size = Pt(24)

    cell = header_table.cell(0, 1)
    cell_paragraph = cell.add_paragraph()
    cell_paragraph.add_run(prod_name).font.size = Pt(12)
    cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT



    ################################################################ Callout section

    

    doc.add_picture(img_path, width=Inches(6.0))
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    ################################################################ At a glance section

    doc.add_paragraph().add_run("At a glance").bold = True

    features = df.loc[df['Tag'].str.endswith('medium'), :]
    for feature in features['4RA85F [Product]']:
        if not pd.isna(feature):
            doc.add_paragraph(feature, style='List Bullet')
    
    footnote_numbers = df[df['ContainerName'].str.endswith('(medium) Footnote Number')]
    
    for footnote in footnote_numbers['4RA85F [Product]']:
        if not pd.isna(footnote):
            doc.add_paragraph(footnote)    
    
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    ################################################################ Operating system section

    ossuppted_values = df.loc[df['Tag'] == 'ossuppted', ['ContainerName', '4RA85F [Product]']].iloc[0]
     
    ossuppted_title = doc.add_heading(ossuppted_values['ContainerName'], level=1)
    ossuppted_title.style.font.size = Pt(14)

    ossuppted_subtitle_replace = ossuppted_values['4RA85F [Product]'].replace('; ', '\n')
    ossuppted_subtitle = doc.add_heading(level=2)
    ossuppted_subtitle.add_run(ossuppted_subtitle_replace).bold = False


            
    doc.save('quickestspecs.docx')

def main():
    loadxlsx()

if __name__ == "__main__":
        main()