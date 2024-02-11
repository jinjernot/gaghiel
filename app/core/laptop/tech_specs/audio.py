from app.core.blocks.paragraph import insertList

def audio_section(doc, html_file, df):
    """Audio techspecs section"""

    # Function to insert the list of values
    insertList(doc, html_file, df, "Audio/Multimedia")
