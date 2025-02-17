from app.core.blocks.paragraph import insert_list

def keyboard_section(doc, df):
    """Keyboard techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, df, "Keyboards/Pointing Devices/Buttons & Function Keys")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)