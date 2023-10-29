
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def audio_section(doc, df):

    audio_paragraph = doc.add_paragraph()
    run = audio_paragraph.add_run("AUDIO / MULTIMEDIA")
    run.font.size = Pt(12)
    run.bold = True
    audio_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    audio_paragraph.add_run().add_break()

    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
