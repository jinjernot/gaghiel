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
from zipfile import ZipFile
import os


def createdocx(file):
    """Create the Quickestspecs"""
    
    # Variables
    doc = Document()
    html_file = 'quickestspecs.html'
    format_document(doc, file, imgs_path)

    # Quickspecs sections
    overview_section(doc, file, html_file, imgs_path)
    tech_specs_section(doc, file, html_file)
    system_unit_section(doc, file, html_file)
    displays_section(doc, file, html_file)
    storage_section(doc, file, html_file)
    network_section(doc, file, html_file)
    audio_section(doc, file, html_file)
    fingerprint_section(doc, file, html_file)
    options_section(doc, file, html_file)
    #superscript(doc, file, html_file)

    docx_file = 'quickestspecs.docx'
    pdf_file = 'quickestspecs.pdf'
    doc.save(docx_file)
    os.rename(docx_file, pdf_file)

    # Create a zip file and add specific files to it
    zip_file_name = 'quickestspecs.zip'
    with ZipFile(zip_file_name, 'w') as zipf:
        zipf.write(html_file)
        zipf.write(pdf_file)
