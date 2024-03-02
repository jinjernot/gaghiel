from app.core.blocks.paragraph import insert_list

def ports_section(doc, html_file, df):
    """Ports techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Ports/Slots")