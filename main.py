import pandas as pd
import glob
from docx import Document
from docx.shared import Pt

def loadxlsx():
    """Load the xlsx file"""
    folder_path = "./xlsx/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")

    for xlsx_file in xlsx_files: #loop through all the files
        createdocx(xlsx_file)

def createdocx(xlsx_file):
    """Create the Quickestspecs"""

    df = pd.read_excel(xlsx_file)
    doc = Document()

    ossuppted_values = df.loc[df['Tag'] == 'ossuppted', ['ContainerName', '4RA85F [Product]']].iloc[0]
     
    ossuppted_title = doc.add_heading(ossuppted_values['ContainerName'], level=1)
    ossuppted_title.style.font.size = Pt(16)

    ossuppted_subtitle = ossuppted_values['4RA85F [Product]'].replace(';', '\n')
    print (ossuppted_subtitle)
    ossuppted_subtitle = doc.add_heading(ossuppted_values['4RA85F [Product]'], level=2)
    ossuppted_subtitle.bold = False

            
    doc.save('quickestspecs.docx')

def main():
    loadxlsx()

if __name__ == "__main__":
        main()