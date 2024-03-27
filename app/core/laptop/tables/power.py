
from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.blocks.table import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK
import pandas as pd

def power_section(doc, file, html_file):
    """Power QS Only Section"""

    # Load xlsx
    #df = pd.read_excel(file, sheet_name='QS-Only Storage')
    df = pd.read_excel(file.stream, sheet_name='QS-Only Power', engine='openpyxl')

    # Add title: Power
    insert_title(doc, "Power", html_file)

    # Add table
    insert_table(doc, df, html_file)

    # Insert HR
    insert_horizontal_line(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
