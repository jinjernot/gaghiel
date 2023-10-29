
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def networking_section(doc, df):

    networking_paragraph = doc.add_paragraph()
    run = networking_paragraph.add_run("NETWORKING/COMMUNICATIONS")
    run.font.size = Pt(12)
    run.bold = True
    networking_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    networking_paragraph.add_run().add_break()

    wlan_subtitle = df.iloc[267, 6]
    wlan_paragraph = doc.add_paragraph()
    run = wlan_paragraph.add_run(wlan_subtitle)
    run.bold = True
    wlan_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    wlan = df.iloc[268:273, 6].tolist()
    wlan = [x for x in wlan if pd.notna(x)]
    wlan_paragraph.add_run().add_break()

    for x in wlan:
        run = wlan_paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    wwlan_subtitle = df.iloc[274, 6]
    wwlan_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = wwlan_paragraph.add_run(wwlan_subtitle)
    run.bold = True
    wwlan_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    wwlan = df.iloc[275:278, 6].tolist()
    wwlan = [x for x in wwlan if pd.notna(x)]
    wwlan_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for x in wwlan:
        run = wwlan_paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    nfc_subtitle = df.iloc[279, 6]
    nfc_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = nfc_paragraph.add_run(nfc_subtitle)
    run.bold = True
    nfc_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    nfc = df.iloc[280:281, 6].tolist()
    nfc = [x for x in nfc if pd.notna(x)]
    nfc_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for x in nfc:
        run = nfc_paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    miracast_subtitle = df.iloc[282, 6]
    miracast_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = miracast_paragraph.add_run(miracast_subtitle)
    run.bold = True
    miracast_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    miracast = df.iloc[283:285, 6].tolist()
    miracast = [x for x in miracast if pd.notna(x)]
    miracast_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for x in miracast:
        run = miracast_paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    ethernet_subtitle = df.iloc[286, 6]
    ethernet_paragraph = doc.add_paragraph()

    # Add the text from the DataFrame to the paragraph
    run = ethernet_paragraph.add_run(ethernet_subtitle)
    run.bold = True
    ethernet_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    ethernet = df.iloc[287:289, 6].tolist()
    ethernet = [x for x in ethernet if pd.notna(x)]
    ethernet_paragraph.add_run().add_break()

     # Add the data from the list to the paragraph
    for x in ethernet:
        run = ethernet_paragraph.add_run(x)
        run.add_break(WD_BREAK.LINE)

    networking_footnotes = df.iloc[291:300, 6].tolist()
    networking_footnotes = [x for x in networking_footnotes if pd.notna(x)]

    # Create a new paragraph
    networking_footnotes_paragraph = doc.add_paragraph()

    # Add the data from the list to the paragraph
    for x in networking_footnotes:
        run = networking_footnotes_paragraph.add_run(x)

        # Set the font color to blue
        run.font.color.rgb = RGBColor(0, 0, 255)  # RGB for blue

        run.add_break(WD_BREAK.LINE)
    insertHR(doc.add_paragraph(), thickness=3)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
