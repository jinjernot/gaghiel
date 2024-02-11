from app.core.blocks.paragraph import insertList

def software_section(doc, html_file, df):
    """Software and security techspecs section"""

    insertList(doc, html_file, df, "Software and Security")
