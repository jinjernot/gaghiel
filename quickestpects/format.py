import pandas as pd
from docx import Document
from docx.shared import Pt

from quickestpects.callouts import callout_section
from quickestpects.at_a_glance import ataglance_section
from quickestpects.operating_system import os_section
from quickestpects.header import header
from quickestpects.footer import footer
from quickestpects.storage_and_drives import hd_section


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
    df = df[df['4RA85F [Product]'] != '##BLANK##'] 

    prod_name = df.loc[df['Tag'] == 'prodname', '4RA85F [Product]'].iloc[0] 
    header(doc, prod_name,)
    footer(doc, imgs_path)
    callout_section(doc, imgs_path,)
    ataglance_section(doc, df)
    os_section(doc, df)
    hd_section(doc)

    

    doc.save('quickestspecs.docx')
