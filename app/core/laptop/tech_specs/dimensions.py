from app.core.blocks.paragraph import insert_list

def dimensions_section(doc, html_file, df):
    """Dimensions techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Weight & Dimensions")