from app.core.blocks.paragraph import insert_list

def audio_section(doc, html_file, df):
    """Audio techspecs section"""

    # Function to insert the list of values
    insert_list(doc, html_file, df, "Audio/Multimedia")
