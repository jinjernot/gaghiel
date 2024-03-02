from app.core.blocks.paragraph import insert_list

def keyboard_section(doc, html_file, df):
    """Keyboard techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Keyboards/Pointing Devices/Buttons & Function Keys")
