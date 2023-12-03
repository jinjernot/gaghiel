from quickestspects.tech_specs.tech_specs import tech_specs_section
from quickestspects.overview.overview import overview_section
from quickestspects.format.format import format_document

from quickestspects.tables.system_unit import system_unit_section

import pandas as pd
from docx import Document



def createdocx(xlsx_file, imgs_path):
    """Create the Quickestspecs"""
    
    # Variables
    doc = Document()
    txt_file = 'quickestspecs.txt'
    format_document(doc, xlsx_file, imgs_path)

    # Quickspecs sections
    #overview_section(doc, xlsx_file, txt_file, df, prod_name, imgs_path)
    tech_specs_section(doc, xlsx_file, txt_file)
    
    #$system_unit_section(doc, xlsx_file)
    
    docx_file = 'quickestspecs.docx'
    doc.save(docx_file)