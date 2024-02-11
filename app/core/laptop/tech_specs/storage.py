from app.core.blocks.paragraph import insertList

def storage_section(doc, html_file, df):
    """Storage techspecs section"""

    insertList(doc, html_file, df, "Storage and Drives")
