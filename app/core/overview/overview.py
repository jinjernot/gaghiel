from app.core.overview.callouts import callout_section
from app.core.overview.at_a_glance import ataglance_section

import pandas as pd

def overview_section(doc, xlsx_file, txt_file, df, prod_name, imgs_path):
    """Add Overview section"""
    
    # Load sheet into df
    df = pd.read_excel(xlsx_file, sheet_name='QS Callouts & Overview')

    # Run the functions to build the overview section
    callout_section(doc,txt_file, df, prod_name, imgs_path)
    ataglance_section(doc, txt_file, df)
