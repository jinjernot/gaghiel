from app.core.blocks.paragraph import insert_list

def audio_section(doc, html_file, df):
    """Audio techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, html_file, df, "Audio/Multimedia")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)