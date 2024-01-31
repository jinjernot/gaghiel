from app.core.laptop.tech_specs.product_name import product_name_section
from app.core.laptop.tech_specs.operating_systems import operating_systems_section
#from app.core.laptop.tech_specs.processors import processors_section
#from app.core.laptop.tech_specs.chipset import chipset_section
from app.core.laptop.tech_specs.graphics import graphics_section
from app.core.laptop.tech_specs.display import display_section
from app.core.laptop.tech_specs.docking import docking_section
from app.core.laptop.tech_specs.storage import storage_section
from app.core.laptop.tech_specs.memory import memory_section
from app.core.laptop.tech_specs.networking import networking_section
from app.core.laptop.tech_specs.audio import audio_section
from app.core.laptop.tech_specs.keyboard import keyboard_section
from app.core.laptop.tech_specs.software import software_section
from app.core.laptop.tech_specs.power import power_section
from app.core.laptop.tech_specs.dimensions import dimensions_section
from app.core.laptop.tech_specs.ports import ports_section
from app.core.laptop.tech_specs.service import service_section


import pandas as pd

def tech_specs_section(doc, file, html_file):
    """TechSpecs Section"""

    # Load sheet into df
    #df = pd.read_excel(file, sheet_name='Tech Specs & QS Features')
    df = pd.read_excel(file.stream, sheet_name='Tech Specs & QS Features', engine='openpyxl')
    
    # Get product name
    prod_name = df.columns[1]
    
    # Run the functions to build the tech specs section
    product_name_section(doc, html_file, prod_name)
    operating_systems_section(doc, html_file, df)
    #processors_section(doc, html_file, df)
    #chipset_section(doc, html_file, df)
    graphics_section(doc, html_file, df)
    display_section(doc, html_file, df)
    docking_section(doc, html_file, df)
    storage_section(doc, html_file, df)
    memory_section(doc, html_file, df)
    networking_section(doc, html_file, df)
    audio_section(doc, html_file, df)
    keyboard_section(doc, html_file, df)
    software_section(doc, html_file, df)
    power_section(doc, html_file, df)
    dimensions_section(doc, html_file, df)
    ports_section(doc, html_file, df)
    service_section(doc, html_file, df)
    