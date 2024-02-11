from app.core.blocks.paragraph import insertList

def networking_section(doc, html_file, df):
    """Network techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Networking /Communications")
