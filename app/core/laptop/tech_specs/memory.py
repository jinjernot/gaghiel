from app.core.blocks.paragraph import insertList

def memory_section(doc, html_file, df):
    """Memory techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Memory")