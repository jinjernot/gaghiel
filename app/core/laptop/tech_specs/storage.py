from app.core.blocks.paragraph import insert_list

def storage_section(doc, html_file, df):
    """Storage techspecs section"""

    insert_list(doc, html_file, df, "Storage and Drives")
