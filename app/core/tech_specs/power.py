from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK

def power_section(doc, txt_file, df):
    """Power techspecs section"""

    # Add the title: POWER
    insertTitle(doc, "POWER", txt_file)

    # Power Supply
    insertSubtitle(doc, txt_file, df, 262, 1)
    insertList(doc, txt_file, df, slice(263, 266),1)

    # Battery
    insertSubtitle(doc, txt_file, df, 266, 1)
    insertList(doc, txt_file, df, slice(267, 269), 1)

    # Battery Recharge Time
    insertSubtitle(doc, txt_file, df, 269, 1)
    insertList(doc, txt_file, df, slice(270, 271), 1)

    # Power Cord
    insertSubtitle(doc, txt_file, df, 271, 1)
    insertList(doc, txt_file, df, slice(272, 273), 1)

    # Battery life
    insertSubtitle(doc, txt_file, df, 273, 1)
    insertList(doc, txt_file, df, slice(274, 279), 1)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(280, 284), 1)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
