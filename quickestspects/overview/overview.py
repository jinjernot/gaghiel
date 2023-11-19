from quickestspects.overview.callouts import callout_section
from quickestspects.overview.at_a_glance import ataglance_section

import pandas as pd

def overview_section(doc, xlsx_file, txt_file, df, prod_name, imgs_path):

    df = pd.read_excel(xlsx_file, sheet_name='QS Callouts & Overview')

    callout_section(doc,txt_file, df, prod_name, imgs_path)
    ataglance_section(doc, txt_file, df)
