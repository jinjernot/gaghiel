from app.core.blocks.paragraph import insertList

def keyboard_section(doc, html_file, df):
    """Keyboard techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Keyboards/Pointing Devices/Buttons & Function Keys")
