from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK


def service_section(doc, html_file, df):
    """Service and support techspecs section"""

    # Add the title: SERVICE AND SUPPORT
    insert_title(doc, "SERVICE AND SUPPORT", html_file)

    # Service and Support
    #insertParagraph(doc, html_file, df, 317, 1)

    # Footnotes
    #insertFootnote(doc, html_file, df, slice(320, 321), 1)

    # HR
    insert_horizontal_line(doc.add_paragraph(), thickness=3)
    insert_html_horizontal_line(html_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
