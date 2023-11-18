import re
from docx import Document

def find_text_in_square_brackets(docx_path):
    doc = Document(docx_path)

    # Regular expression to match text inside square brackets (allowing numbers)
    pattern = r'\[(\d+)\]'

    # Find all matches in the document
    matches = []

    # Process paragraphs
    for paragraph in doc.paragraphs:
        matches.extend(re.findall(pattern, paragraph.text))

    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches.extend(re.findall(pattern, cell.text))

    return matches

def superscript_numbers_in_square_brackets(docx_path):
    doc = Document(docx_path)

    # Find numbers in square brackets
    matches = find_text_in_square_brackets(docx_path)

    # Loop through paragraphs
    for paragraph in doc.paragraphs:
        for match in matches:
            # Superscript the numbers inside square brackets
            paragraph.text = re.sub(rf'\[({re.escape(match)})\]', rf'[sub]{match}[/sub]', paragraph.text, flags=re.IGNORECASE)

    # Loop through tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for match in matches:
                    # Superscript the numbers inside square brackets
                    cell.text = re.sub(rf'\[({re.escape(match)})\]', rf'[sub]{match}[/sub]', cell.text, flags=re.IGNORECASE)

    # Save the modified document
    doc.save('output.docx')

# Example usage
docx_path = 'quickestspecs.docx'
superscript_numbers_in_square_brackets(docx_path)
