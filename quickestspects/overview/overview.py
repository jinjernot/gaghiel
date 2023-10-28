import pandas as pd

from quickestspects.overview.callouts import callout_section
from quickestspects.overview.at_a_glance import ataglance_section


def overview_section(xlsx_file, doc, df, prod_name, imgs_path):

    df = pd.read_excel(xlsx_file, sheet_name='QS Callouts & Overview')

    callout_section(doc, df, prod_name, imgs_path)
    ataglance_section(doc, df)
