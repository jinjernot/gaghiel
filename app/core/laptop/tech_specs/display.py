from app.core.blocks.paragraph import insert_list

def display_section(doc, html_file, df):
    """Display techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Display")
