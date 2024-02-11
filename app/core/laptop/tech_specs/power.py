from app.core.blocks.paragraph import insertList

def power_section(doc, html_file, df):
    """Power techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Power")