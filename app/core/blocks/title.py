from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt

def insertTitle(doc, title, html_file):
    # Add the title to the Word document
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(title)
    run.font.size = Pt(12)
    run.bold = True
    #paragraph.add_run().add_break()

    # Write the HTML title
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(f"<b><h1>{title}</h1></b>\n")

def insertSubtitle(doc, html_file, df,  iloc_row, iloc_column):
    # Add the subtitle to the Word document
    paragraph = doc.add_paragraph()
    subtitle = df.iloc[iloc_row, iloc_column]
    run = paragraph.add_run(subtitle)
    run.bold = True
    #run.add_break(WD_BREAK.LINE)

    # Write the HTML subtitle
    with open(html_file, 'a', encoding='utf-8') as txt:
        txt.write(f"<b><p>{subtitle}</p></b>\n")
