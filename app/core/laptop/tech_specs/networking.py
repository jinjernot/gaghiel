from app.core.blocks.paragraph import insert_list

def networking_section(doc, html_file, df):
    """Network techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Networking /Communications")
