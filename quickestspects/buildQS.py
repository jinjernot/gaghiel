import pandas as pd
from docx import Document
from docx.shared import Pt

from quickestspects.callouts import callout_section
from quickestspects.at_a_glance import ataglance_section
from quickestspects.panel_specifications import panel_specs
from quickestspects.header import header
from quickestspects.footer import footer



def set_margins(doc):
    sections = doc.sections
    for section in sections:
        section.left_margin = Pt(20)  
        section.right_margin = Pt(20)  
        section.top_margin = Pt(20)  
        section.bottom_margin = Pt(20)  


def createdocx(xlsx_file, imgs_path):
    """Create the Quickestspecs"""
    doc = Document()
    set_margins(doc)  
    df = pd.read_excel(xlsx_file) 
    prod_name = df['ItemName'].iloc[0]
    header(doc, prod_name)
    footer(doc, imgs_path)
    callout_section(df, doc, imgs_path)
    panel_specs(doc, df)

    # Save as DOCX
    docx_file = 'quickestspecs.docx'
    doc.save(docx_file)