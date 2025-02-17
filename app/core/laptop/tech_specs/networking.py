from app.core.blocks.paragraph import insert_list

def networking_section(doc, df):
    """Network techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, df, "Networking /Communications")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)