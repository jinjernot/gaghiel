from app.core.blocks.paragraph import insertList

def display_section(doc, html_file, df):
    """Display techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Display")
