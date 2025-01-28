from app.core.blocks.paragraph import insert_list

def dimensions_section(doc, df):
    """Dimensions techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, df, "Weight & Dimensions")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)