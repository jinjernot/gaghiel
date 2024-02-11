from app.core.blocks.paragraph import insertList

def ports_section(doc, html_file, df):
    """Ports techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Ports/Slots")