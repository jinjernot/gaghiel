from app.core.overview.callouts import callout_section
from app.core.overview.at_a_glance import ataglance_section

import pandas as pd

def overview_section(doc, file, html_file, imgs_path):
    """Add Overview section"""
    
    # Load sheet into df
    df = pd.read_excel(file, sheet_name='Callouts')
    #df = pd.read_excel(file.stream, sheet_name='Tech Specs & QS Features', engine='openpyxl')

    prod_name = df.columns[1]
    # Run the functions to build the overview section
    callout_section(doc, html_file, prod_name, imgs_path, df)
    #ataglance_section(doc, html_file, df)
