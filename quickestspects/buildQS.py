from quickestspects.tech_specs.tech_specs import tech_specs_section
from quickestspects.overview.overview import overview_section
from quickestspects.format.format import format_document

import pandas as pd
from docx import Document



def createdocx(xlsx_file, imgs_path):
    """Create the Quickestspecs"""

    #get product name
    df = pd.read_excel(xlsx_file,sheet_name = 'Metadata') 
    prod_name = df.columns[1]
    
    doc = Document()

    format_document(doc, prod_name, imgs_path)
    overview_section(xlsx_file, doc, df, prod_name, imgs_path)
    tech_specs_section(xlsx_file, doc, df, prod_name)

    # Save as DOCX
    docx_file = 'quickestspecs.docx'
    doc.save(docx_file)