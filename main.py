from flask import Flask, render_template
from app.buildDocx import createdocx
import glob

app = Flask(__name__)

@app.route('/quickspecs')
def index():
    """Homepage with a button to generate DOCX files"""
    return render_template('index.html')

@app.route('/quickspecs/generate_docx')
def generate_docx():
    """Endpoint to generate DOCX files"""
    folder_path = "./xlsx/"
    imgs_path = "./imgs/"
    xlsx_files = glob.glob(folder_path + "*.xlsx")

    for xlsx_file in xlsx_files:
        createdocx(xlsx_file, imgs_path)

    return "DOCX files generated successfully!"

if __name__ == "__main__":
    app.run(debug=True)
