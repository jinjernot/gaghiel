from quickestspects.blocks.paragraph import *
from quickestspects.blocks.title import *
from quickestspects.format.hr import *

from docx.enum.text import WD_BREAK

import pandas as pd

def display_section(doc, txt_file, df):
    """Display techspecs section"""

    # Add the title: DISPLAY
    insertTitle(doc, "DISPLAY", txt_file)   

    # Non-Touch
    insertSubtitle(doc, txt_file, df, 130, 6)
    insertList(doc, txt_file, df, slice(131, 142), 6)

    # Touch
    insertSubtitle(doc, txt_file, df, 145, 6)
    insertList(doc, txt_file, df, slice(146, 149), 6)

    # Display Port
    insertSubtitle(doc, txt_file, df, 155, 6)
    insertList(doc, txt_file, df, slice(156, 157), 6)

    # Display Support
    insertSubtitle(doc, txt_file, df, 158, 6)
    insertList(doc, txt_file, df, slice(159, 161), 6)

    # Display Size
    insertSubtitle(doc, txt_file, df, 161, 6)
    insertList(doc, txt_file, df, slice(162, 164), 6)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(186, 191), 6)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)


    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
