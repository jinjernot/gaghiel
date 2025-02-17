from app.core.blocks.paragraph import insert_list

def operating_systems_section(doc, df):
    """Operating system techspecs section"""
  
    try:
        # Function to insert the list of values
        insert_list(doc, df, "Operating System")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)