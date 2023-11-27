from quickestspects.blocks.paragraph import *
from quickestspects.blocks.title import *
from quickestspects.format.hr import *

from docx.enum.text import WD_BREAK

def networking_section(doc, txt_file, df):
    """Network techspecs section"""

    # Add the title: NETWORKINGS
    insertTitle(doc, "NETWORKING", txt_file)

    # Wlan
    insertSubtitle(doc, txt_file, df, 267, 6)
    insertList(doc, txt_file, df, slice(268, 273), 6)

    # Wwlan
    insertSubtitle(doc, txt_file, df, 274, 6)
    insertList(doc, txt_file, df, slice(275, 278), 6)

    # NFC
    insertSubtitle(doc, txt_file, df, 279, 6)
    insertList(doc, txt_file, df, slice(280, 281), 6)

    # Miracast
    insertSubtitle(doc, txt_file, df, 282, 6)
    insertList(doc, txt_file, df, slice(283, 285), 6)

    # Ethernet
    insertSubtitle(doc, txt_file, df, 286, 6)
    insertList(doc, txt_file, df, slice(287, 289), 6)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(291, 300), 6)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)