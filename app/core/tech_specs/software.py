from app.core.blocks.paragraph import *
from app.core.blocks.title import *
from app.core.format.hr import *

from docx.enum.text import WD_BREAK

def software_section(doc, txt_file, df):
    """Software and security techspecs section"""

    # Add the title: SOFTWARE AND SECURITY
    insertTitle(doc, "SOFTWARE AND SECURITY", txt_file)

    # Software
    insertSubtitle(doc, txt_file, df, 175, 1)
    insertList(doc, txt_file, df, slice(176, 196),1)

    # Manageability Features
    insertSubtitle(doc, txt_file, df, 196, 1)
    insertList(doc, txt_file, df, slice(197, 205), 1)

    # Security Management
    insertSubtitle(doc, txt_file, df, 205, 1)
    insertList(doc, txt_file, df, slice(206, 217), 1)

    # Security- TPM
    insertSubtitle(doc, txt_file, df, 217, 1)
    insertList(doc, txt_file, df, slice(218, 219), 1)

    # TCG TPM 2.0
    insertSubtitle(doc, txt_file, df, 219, 1)
    insertList(doc, txt_file, df, slice(220, 221), 1)

    # FIPS 140-2 Compliant: Yes
    insertSubtitle(doc, txt_file, df, 221, 1)
    insertList(doc, txt_file, df, slice(223, 223), 1)

    # BIOS
    insertSubtitle(doc, txt_file, df, 223, 1)
    insertList(doc, txt_file, df, slice(224, 231), 1)

    # Smartcard Reader
    insertSubtitle(doc, txt_file, df, 231, 1)
    insertList(doc, txt_file, df, slice(232, 235), 1)

    # IPv6 Support
    insertSubtitle(doc, txt_file, df, 235, 1)
    insertList(doc, txt_file, df, slice(236, 237), 1)

    # FirstNet Certified
    insertSubtitle(doc, txt_file, df, 237, 1)
    insertList(doc, txt_file, df, slice(238, 242), 1)

    # Footnotes
    insertFootnote(doc, txt_file, df, slice(243, 260), 1)


    # HR
    insertHR(doc.add_paragraph(), thickness=3)
    insertHTMLhr(txt_file)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
