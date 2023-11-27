from quickestspects.format.hr import *
from quickestspects.blocks.title import  insertTitle
from quickestspects.blocks.paragraph import insertParagraph

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def chipset_section(doc, txt_file, df):

    # Add the title
    insertTitle(doc, "CHIPSET", txt_file)
    
    # Add paragraph
    insertParagraph(doc, txt_file, df, 90, 6)

    # Add HR
    insertHR(doc.add_paragraph(), thickness=3)

    # Add HTML <hr>
    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')
