from flask import Flask, request, render_template, send_from_directory
from app.core.laptop.build_laptop import createdocx
import config

app = Flask(__name__)
app.use_static_for = 'static'
# Configuration
app.config.from_object(config)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['VALID_FILE_EXTENSIONS']


@app.route('/app3')
def index():
    """Homepage with a button to generate DOCX files"""
    return render_template('index.html')    

@app.route('/quickspecs/generate_docx', methods=['POST'])
def generate_docx():
    
    if 'MAT' in request.files:
        file = request.files['MAT']
        try:
            if allowed_file(file.filename):
                createdocx(file)
                return send_from_directory('.', 'quickspecs.zip', as_attachment=True)
        except Exception as e:
            print(e)
            return render_template('error.html'), 500
    return render_template('error.html'), 400

if __name__ == "__main__":
    app.run(debug=True)