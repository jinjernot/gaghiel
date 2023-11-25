from quickestspects.format.hr import insertHR

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt, RGBColor
import pandas as pd

def display_section(doc, txt_file, df):

    paragraph = doc.add_paragraph()
    run = paragraph.add_run("DISPLAY")
    run.font.size = Pt(12)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write("<h1><b>DISPLAY</h1></b>\n")

    non_touch_subtitle = df.iloc[130, 6]
    non_touch_paragraph = doc.add_paragraph()
    run = non_touch_paragraph.add_run(non_touch_subtitle)
    run.bold = True
    non_touch_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    non_touch_paragraph.add_run().add_break()
    non_touch = df.iloc[131:142, 6].tolist()
    non_touch = [disp for disp in non_touch if pd.notna(disp)]

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{non_touch_subtitle}</p>\n")

    for disp in non_touch:
        run = non_touch_paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{disp}</p>\n")

    touch_subtitle = df.iloc[145, 6]
    paragraph = doc.add_paragraph()

    run = paragraph.add_run(touch_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{touch_subtitle}</p>\n")

    touch = df.iloc[146:149, 6].tolist()
    touch = [disp for disp in touch if pd.notna(disp)]
    paragraph.add_run().add_break()

    for disp in touch:
        run = paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{disp}</p>\n")

    displayport_subtitle = df.iloc[155, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(displayport_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{displayport_subtitle}</p>\n")


    displayport = df.iloc[156:157, 6].tolist()
    displayport = [disp for disp in displayport if pd.notna(disp)]

    for disp in displayport:
        run = paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)

        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{disp}</p>\n")

    display_support_subtitle = df.iloc[158, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(display_support_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{display_support_subtitle}</p>\n")

    display_support = df.iloc[159:161, 6].tolist()
    display_support = [disp for disp in display_support if pd.notna(disp)]

    for disp in display_support:
        run = paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)
        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{disp}</p>\n")

    display_size_subtitle = df.iloc[161, 6]
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(display_size_subtitle)
    run.bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.add_run().add_break()

    with open(txt_file, 'a') as txt:
        txt.write(f"<p>{display_size_subtitle}</p>\n")

    display_size = df.iloc[162:164, 6].tolist()
    display_size = [disp for disp in display_size if pd.notna(disp)]

    for disp in display_size:
        run = paragraph.add_run(disp)
        run.add_break(WD_BREAK.LINE)
        with open(txt_file, 'a') as txt:
            txt.write(f"<p>{disp}</p>\n")

    display_footnotes = df.iloc[186:191, 6].tolist()
    display_footnotes = [disp_footnote for disp_footnote in display_footnotes if pd.notna(disp_footnote)]

    paragraph = doc.add_paragraph()

    for disp_footnote in display_footnotes:
        run = paragraph.add_run(disp_footnote)
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.add_break(WD_BREAK.LINE)

    html_footnotes = '<div style="color: blue;">\n'
    for disp_footnote in display_footnotes:
        html_footnotes += f'  <span>{disp_footnote}</span>\n'
    html_footnotes += '</div>\n'
    with open(txt_file, 'a') as txt:
            txt.write(html_footnotes)

    insertHR(doc.add_paragraph(), thickness=3)

    with open(txt_file, 'a') as txt:
        txt.write('<hr align="center" SIZE="2" width="100%">\n')

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
