from app.core.blocks.paragraph import insertList

def graphics_section(doc, html_file, df):
    """Graphics techspecs section"""
    
    # Function to insert the list of values
    insertList(doc, html_file, df, "Graphics")
