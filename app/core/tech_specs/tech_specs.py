from app.core.tech_specs.product_name import product_name_section
from app.core.tech_specs.operating_systems import operating_systems_section
from app.core.tech_specs.processors import processors_section
from app.core.tech_specs.chipset import chipset_section
from app.core.tech_specs.graphics import graphics_section
from app.core.tech_specs.display import display_section
from app.core.tech_specs.docking import docking_section
from app.core.tech_specs.storage import storage_section
from app.core.tech_specs.memory import memory_section
from app.core.tech_specs.networking import networking_section
from app.core.tech_specs.audio import audio_section
from app.core.tech_specs.keyboard import keyboard_section
from app.core.tech_specs.software import software_section
from app.core.tech_specs.power import power_section
from app.core.tech_specs.dimensions import dimensions_section
from app.core.tech_specs.ports import ports_section
from app.core.tech_specs.service import service_section


import pandas as pd

def tech_specs_section(doc, file, txt_file):
    """TechSpecs Sections"""

    # Load sheet into df
    df = pd.read_excel(file, sheet_name='Tech Specs & QS Features')
    prod_name = df.columns[1]
    
    # Run the functions to build the tech specs section
    product_name_section(doc, txt_file, prod_name)
    operating_systems_section(doc, txt_file, df)
    #processors_section(doc, txt_file, df)
    #chipset_section(doc, txt_file, df)
    graphics_section(doc, txt_file, df)
    display_section(doc, txt_file, df)
    docking_section(doc, txt_file, df)
    storage_section(doc, txt_file, df)
    memory_section(doc, txt_file, df)
    networking_section(doc, txt_file, df)
    audio_section(doc, txt_file, df)
    keyboard_section(doc, txt_file, df)
    software_section(doc, txt_file, df)
    power_section(doc, txt_file, df)
    dimensions_section(doc, txt_file, df)
    ports_section(doc, txt_file, df)
    service_section(doc, txt_file, df)
    