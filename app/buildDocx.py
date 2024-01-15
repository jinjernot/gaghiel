from app.core.tech_specs.tech_specs import tech_specs_section
from app.core.overview.overview import overview_section
from app.core.format.format import format_document

from app.core.tables.system_unit import system_unit_section
from app.core.tables.displays import displays_section
from app.core.tables.options import options_section
from app.core.tables.audio import audio_section
from app.core.tables.fingerprint import fingerprint_section
from app.core.tables.storage import storage_section
from app.core.tables.network import network_section
#from app.core.format.superscript import superscript

import pandas as pd
from docx import Document
imgs_path = "./imgs/"


def createdocx(file):
    """Create the Quickestspecs"""
    
    # Variables
    doc = Document()
    txt_file = 'quickestspecs.txt'
    format_document(doc, file, imgs_path)

    # Quickspecs sections
    overview_section(doc, file, txt_file, imgs_path)
    tech_specs_section(doc, file, txt_file)
    system_unit_section(doc, file, txt_file)
    displays_section(doc, file, txt_file)
    storage_section(doc, file, txt_file)
    network_section(doc, file, txt_file)
    audio_section(doc, file, txt_file)
    fingerprint_section(doc, file, txt_file)
    options_section(doc, file, txt_file)
    #superscript(doc, file, txt_file)

    
    docx_file = 'quickestspecs.docx'
    doc.save(docx_file)