from quickestspects.format.hr import *
from quickestspects.blocks.title import  insertTitle

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def networking_section(doc, txt_file, df):

    insertTitle(doc, "NETWORKING", txt_file)


    wlan_subtitle = df.iloc[267, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(wlan_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    wlan = df.iloc[268:273, 6].tolist()
    wlan = [x for x in wlan if pd.notna(x)]
    paragraph.add_run().add_break()

    for x in wlan:
        run = paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    wwlan_subtitle = df.iloc[274, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(wwlan_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    wwlan = df.iloc[275:278, 6].tolist()
    wwlan = [x for x in wwlan if pd.notna(x)]
    paragraph.add_run().add_break()

    for x in wwlan:
        run = paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    nfc_subtitle = df.iloc[279, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(nfc_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    nfc = df.iloc[280:281, 6].tolist()
    nfc = [x for x in nfc if pd.notna(x)]
    paragraph.add_run().add_break()

    for x in nfc:
        run = paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    miracast_subtitle = df.iloc[282, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(miracast_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    miracast = df.iloc[283:285, 6].tolist()
    miracast = [x for x in miracast if pd.notna(x)]
    paragraph.add_run().add_break()

    for x in miracast:
        run = paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    ethernet_subtitle = df.iloc[286, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(ethernet_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    ethernet = df.iloc[287:289, 6].tolist()
    ethernet = [x for x in ethernet if pd.notna(x)]
    paragraph.add_run().add_break()

    for x in ethernet:
        run = paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    networking_footnotes = df.iloc[291:300, 6].tolist()
    networking_footnotes = [x for x in networking_footnotes if pd.notna(x)]

    paragraph = doc.add_paragraph()

    for x in networking_footnotes:
        run = paragraph.add_run(x)

        run.font.color.rgb = RGBColor(0, 0, 255) 

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
