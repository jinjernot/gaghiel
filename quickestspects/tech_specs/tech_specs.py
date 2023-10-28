import pandas as pd

#techspecs section
from quickestspects.tech_specs.product_name import product_name_section
from quickestspects.tech_specs.operating_systems import operating_systems_section
from quickestspects.tech_specs.processors import processors_section
from quickestspects.tech_specs.chipset import chipset_section
from quickestspects.tech_specs.graphics import graphics_section
from quickestspects.tech_specs.display import display_section


def tech_specs_section(xlsx_file, doc, df, prod_name):

    df = pd.read_excel(xlsx_file, sheet_name='Tech Specs & QS Features')
    
    product_name_section(doc, prod_name)
    operating_systems_section(doc, df)
    processors_section(doc, df)
    chipset_section(doc, df)
    graphics_section(doc, df)
    display_section(doc, df)