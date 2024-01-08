from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK

def dimensions_section(doc, txt_file, df):
    """Dimensions techspecs section"""

    # Add title
    insertTitle(doc, "WEIGHTS & DIMENSIONS", txt_file)

    # Product Weight
    insertSubtitle(doc, txt_file, df, 286, 1)
    insertList(doc, txt_file, df, slice(287, 288),1)

    # Product Dimensions (w x d x h)
    insertSubtitle(doc, txt_file, df, 288, 1)
    insertList(doc, txt_file, df, slice(289, 291), 1)

    # Package Dimensions (w x d x h)
    insertSubtitle(doc, txt_file, df, 291, 1)
    insertList(doc, txt_file, df, slice(292, 293), 1)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(295, 297), 1)

    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)