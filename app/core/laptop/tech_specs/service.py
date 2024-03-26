from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK


def service_section(doc, html_file, df):
    """Service and support techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Service and Support")