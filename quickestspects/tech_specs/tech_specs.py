from quickestspects.tech_specs.product_name import product_name_section
from quickestspects.tech_specs.operating_systems import operating_systems_section
from quickestspects.tech_specs.processors import processors_section
from quickestspects.tech_specs.chipset import chipset_section
from quickestspects.tech_specs.graphics import graphics_section
from quickestspects.tech_specs.display import display_section
from quickestspects.tech_specs.storage import storage_section
from quickestspects.tech_specs.memory import memory_section
from quickestspects.tech_specs.networking import networking_section

import pandas as pd

def tech_specs_section(doc, xlsx_file, txt_file, df, prod_name):

    df = pd.read_excel(xlsx_file, sheet_name='Tech Specs & QS Features')
    
    product_name_section(doc, txt_file, prod_name)
    operating_systems_section(doc, txt_file, df)
    processors_section(doc, df)
    chipset_section(doc, df)
    graphics_section(doc, df)
    display_section(doc, df)
    storage_section(doc, df)
    memory_section(doc, df)
    networking_section(doc, df)