from app.core.blocks.paragraph import insert_list

def software_section(doc, html_file, df):
    """Software and security techspecs section"""

    insert_list(doc, html_file, df, "Software and Security")
