from quickestspects.blocks.paragraph import *
from quickestspects.blocks.title import *
from quickestspects.format.hr import *

from docx.enum.text import WD_BREAK

def memory_section(doc, txt_file, df):
    """Memory tectspecs section"""

    # Add the title: MEMORY
    insertTitle(doc, "MEMORY", txt_file)

    # Maximum memory
    insertSubtitle(doc, txt_file, df, 235, 6)
    insertParagraph(doc, txt_file, df, 236, 6)

    # Primary memory
    insertSubtitle(doc, txt_file, df, 237, 6)
    insertList(doc, txt_file, df, slice(238, 246), 6)

    # Memory slots
    insertSubtitle(doc, txt_file, df, 247, 6)
    insertList(doc, txt_file, df, slice(248, 255), 6)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(257, 258), 6)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)