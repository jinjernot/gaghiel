
from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import RGBColor
from docx.shared import Pt
import pandas as pd

def graphics_section(doc, txt_file, df):

    graphics_paragraph = doc.add_paragraph()
    run = graphics_paragraph.add_run("GRAPHICS")
    run.font.size = Pt(12)
    run.bold = True
    graphics_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    graphics_paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write("<h1>GRAPHICS</h1>\n")

    integrated_subtitle = df.iloc[102, 6]
    integrated_paragraph = doc.add_paragraph()
    run = integrated_paragraph.add_run(integrated_subtitle)
    run.bold = True
    integrated_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    integrated_paragraph.add_run().add_break()

    integrated = df.iloc[103:108, 6].tolist()
    integrated = [gfx for gfx in integrated if pd.notna(gfx)]
    
    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{integrated_subtitle}</p>\n")

    for gfx in integrated:
        run = integrated_paragraph.add_run(gfx)

        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{gfx}</p>\n")

    discrete_subtitle = df.iloc[108, 6]
    discrete_paragraph = doc.add_paragraph()

    run = discrete_paragraph.add_run(discrete_subtitle)
    run.bold = True
    discrete_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    integrated_paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{discrete_subtitle}</p>\n")

    discrete = df.iloc[110:111, 6].tolist()
    discrete = [gfx for gfx in discrete if pd.notna(gfx)]
    discrete_paragraph.add_run().add_break()

    for gfx in discrete:
        run = discrete_paragraph.add_run(gfx)

        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{gfx}</p>\n")
            print(gfx)

    supports_subtitle = df.iloc[111, 6]
    supports_paragraph = doc.add_paragraph()
    run = supports_paragraph.add_run(supports_subtitle)
    run.bold = True
    supports_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{supports_subtitle}</p>\n")

    supports = df.iloc[112:116, 6].tolist()
    supports = [gfx for gfx in supports if pd.notna(gfx)]
    supports_paragraph.add_run().add_break()

    for gfx in supports:
        run = supports_paragraph.add_run(gfx)
        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{gfx}</p>\n")
            print(gfx)

    graphics_footnotes = df.iloc[117:121, 6].tolist()
    graphics_footnotes = [gfx_footnote for gfx_footnote in graphics_footnotes if pd.notna(gfx_footnote)]
    graphics_footnote_paragraph = doc.add_paragraph()

    for gfx_footnote in graphics_footnotes:
        run = graphics_footnote_paragraph.add_run(gfx_footnote)
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.add_break(WD_BREAK.LINE)

    html_footnotes = '<div style="color: blue;">\n'
    for gfx_footnote in graphics_footnotes:
        html_footnotes += f'  <span>{gfx_footnote}</span>\n'
    html_footnotes += '</div>\n'
    with open(txt_file, 'a') as txt:
            txt.write(html_footnotes)

    insertHR(doc.add_paragraph(), thickness=3)

    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
