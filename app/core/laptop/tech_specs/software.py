from app.core.blocks.paragraph import insert_list

def software_section(doc, html_file, df):
    """Software and security techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, html_file, df, "Software and Security")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)