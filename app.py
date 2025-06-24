from flask import Flask, request, render_template_string, send_file
import os
import pandas as pd
from generate_report_aub2connect import generate_ppt_report

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template_string('''
        <h2>Upload Excel File</h2>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file" required>
            <button type="submit">Upload</button>
        </form>
    ''')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file:
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        output_path = os.path.join(OUTPUT_FOLDER, 'result.pptx')
        logo_path = 'logo.png'

        file.save(input_path)
        generate_ppt_report(input_path, output_path, logo_path)
        return send_file(output_path, as_attachment=True)
    return 'No file uploaded'

if __name__ == '__main__':
    app.run(debug=True)
