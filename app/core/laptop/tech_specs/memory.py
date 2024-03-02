from app.core.blocks.paragraph import insert_list

def memory_section(doc, html_file, df):
    """Memory techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Memory")