from quickestspects.format.header import header
from quickestspects.format.footer import footer

from docx.shared import Pt

def set_margins(doc):
    """Set document margins"""
    
    sections = doc.sections
    for section in sections:
        section.left_margin = Pt(20)  
        section.right_margin = Pt(20)  
        section.top_margin = Pt(20)  
        section.bottom_margin = Pt(20)  

def default_font(doc):
    """Set default font"""
    styles = doc.styles
    default_style = styles['Normal']
    font = default_style.font
    font.name = 'HP Simplified'
    font.size = Pt(10)


def format_document(doc, prod_name, imgs_path):
    """Apply formatting to document"""
    header(doc, prod_name)
    footer(doc, imgs_path)
    set_margins(doc)
    default_font(doc)