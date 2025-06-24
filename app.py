import os
from flask import Flask, request, send_file, render_template_string
from generate_report_aub2connect import generate_ppt_report

# 确保 uploads 和 output 文件夹存在
for folder in ['uploads', 'output']:
    if not os.path.exists(folder):
        os.makedirs(folder)

app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string("""
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload Excel to Generate PPT</h1>
    <form method=post enctype=multipart/form-data action="/upload">
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    """)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if not file:
        return 'No file uploaded.'

    input_path = os.path.join('uploads', file.filename)
    output_path = os.path.join('output', file.filename.replace('.xlsx', '.pptx'))
    logo_path = 'logo.png'

    file.save(input_path)

    try:
        generate_ppt_report(input_path, output_path, logo_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
