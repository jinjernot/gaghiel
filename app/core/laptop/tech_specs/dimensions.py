from app.core.blocks.paragraph import insertList

def dimensions_section(doc, html_file, df):
    """Dimensions techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Weight & Dimensions")