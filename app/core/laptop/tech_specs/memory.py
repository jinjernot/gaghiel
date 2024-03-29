from app.core.blocks.paragraph import insert_list

def memory_section(doc, html_file, df):
    """Memory techspecs section"""

    try:
        # Function to insert the list of values
        insert_list(doc, html_file, df, "Memory")
    except Exception as e:
        print(f"An error occurred: {e}")
        return str(e)