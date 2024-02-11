from app.core.format.header import header
from app.core.format.footer import footer
import json

from docx.shared import Pt

def read_bold_words_from_json(json_file):
    with open(json_file, 'r') as f:
        data = json.load(f)
        return data.get('bold_words', [])

def set_margins(doc, file):
    """Set document margins"""

    sections = doc.sections
    for section in sections:
        section.left_margin = Pt(20)  
        section.right_margin = Pt(20)  
        section.top_margin = Pt(20)  
        section.bottom_margin = Pt(20)  

def default_font(doc):
    """Set default font"""

    styles = doc.styles
    default_style = styles['Normal']
    font = default_style.font
    font.name = 'HP Simplified'
    font.size = Pt(10)

def bold_font(doc):
    #bold_words = read_bold_words_from_json('app/core/format/bold_words.json')
    bold_words = read_bold_words_from_json('/home/garciagi/qs/app/core/format/bold_words.json')
    
    # Iterate through the paragraphs in the document
    for paragraph in doc.paragraphs:
        # Iterate through the runs in each paragraph
        for run in paragraph.runs:
            # Iterate through the bold words list
            for word in bold_words:
                # If the word is found in the run text
                if word in run.text:
                    # Find the index of the word in the run text
                    index = run.text.find(word)
                    if index != -1:
                        # Check if the word is at the beginning of the line
                        if index == 0 or run.text[index - 1] == '\n':
                            run.bold = True

def format_document(doc, file, imgs_path):
    """Apply formatting to document"""

    header(doc, file)
    footer(doc, imgs_path)
    set_margins(doc, file)
    default_font(doc)
    bold_font(doc)