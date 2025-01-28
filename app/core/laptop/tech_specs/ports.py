from app.core.blocks.paragraph import insert_list

def ports_section(doc, df):
    """Ports techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, df, "Ports/Slots")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)