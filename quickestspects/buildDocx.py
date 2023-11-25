from quickestspects.tech_specs.tech_specs import tech_specs_section
from quickestspects.overview.overview import overview_section
from quickestspects.format.format import format_document

from quickestspects.tables.system_unit import system_unit_section
#from quickestspects.format.superscript import process_superscript

import pandas as pd
from docx import Document



def createdocx(xlsx_file, imgs_path, txt_file):
    """Create the Quickestspecs"""

    df = pd.read_excel(xlsx_file,sheet_name = 'Metadata') 
    prod_name = df.columns[1]
    
    doc = Document()

    
    format_document(doc, prod_name, imgs_path)

    # Quickspecs sections
    overview_section(doc, xlsx_file, txt_file, df, prod_name, imgs_path)
    tech_specs_section(doc, xlsx_file, txt_file, df, prod_name)
    
    system_unit_section(doc, xlsx_file, df)

    #process_superscript(doc)
    
    docx_file = 'quickestspecs.docx'
    doc.save(docx_file)