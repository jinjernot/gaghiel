from app.core.blocks.paragraph import insertList

def operating_systems_section(doc, html_file, df):
    """Operating system techspecs section"""
  
    # Function to insert the list of values
    insertList(doc, html_file, df, "Operating Systems")
