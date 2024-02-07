from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK

def graphics_section(doc, html_file, df):
    """Graphics techspecs section"""
    
    # Add the title: GRAPHICS
    insertTitle(doc, "GRAPHICS", html_file)

    # Integrated
    #insertSubtitle(doc, html_file, df, 13, 1)
    insertList(doc, html_file, df, "Integrated")

    # Discrete
    # insertSubtitle(doc, html_file, df, 16, 1)
    # insertList(doc, html_file, df, slice(17, 18), 1)

    # Supports
    # insertSubtitle(doc, html_file, df, 18, 1)
    # insertList(doc, html_file, df, slice(19, 21), 1)

    # Footnotes
    # insertFootnote(doc, html_file, df, slice(22, 24), 1)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(html_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
