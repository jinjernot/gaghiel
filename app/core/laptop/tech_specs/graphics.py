from app.core.blocks.paragraph import insert_list

def graphics_section(doc, html_file, df):
    """Graphics techspecs section"""
    
    # Function to insert the list of values
    insert_list(doc, html_file, df, "Graphics")
