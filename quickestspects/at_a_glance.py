import pandas as pd
from docx.enum.text import WD_BREAK

def ataglance_section(doc, df):

    doc.add_paragraph().add_run("At a Glance").bold = True

    features = df.loc[df['Tag'].str.endswith('medium'), :]
    for feature in features['4RA85F [Product]']:
        if not pd.isna(feature):
            doc.add_paragraph(feature, style='List Bullet')
    
    footnote_numbers = df[df['ContainerName'].str.endswith('(medium) Footnote Number')]
    
    for footnote in footnote_numbers['4RA85F [Product]']:
        if not pd.isna(footnote):
            doc.add_paragraph(footnote)    
    
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)