import os
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # ✅ 避免 Render 出现 font cache 问题
import matplotlib.pyplot as plt
from flask import Flask, request, render_template_string, send_file
from pptx import Presentation
from pptx.util import Inches

UPLOAD_FOLDER = 'uploads'
CHART_FOLDER = 'charts'
OUTPUT_FOLDER = 'output'

for folder in [UPLOAD_FOLDER, CHART_FOLDER, OUTPUT_FOLDER]:
    if os.path.exists(folder) and not os.path.isdir(folder):
        os.remove(folder)
    os.makedirs(folder, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

UPLOAD_FORM = """
<!doctype html>
<title>Upload Excel</title>
<h1>Upload Excel to Generate Report</h1>
<form action="/upload" method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Upload>
</form>
"""

@app.route('/')
def index():
    return render_template_string(UPLOAD_FORM)

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    output_path = os.path.join(OUTPUT_FOLDER, 'report.pptx')
    generate_ppt_report(filepath, output_path)
    return send_file(output_path, as_attachment=True)

def generate_ppt_report(excel_path, output_path):
    df = pd.read_excel(excel_path)
    checked_in = df['check-in'].str.upper().value_counts().get('Y', 0)
    not_checked_in = df['check-in'].str.upper().value_counts().get('N', 0)
    attendance_rate = round((checked_in / len(df)) * 100, 2)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # 饼图
    fig, ax = plt.subplots()
    ax.pie([checked_in, not_checked_in], labels=['Checked In', 'Not Checked'], autopct='%1.1f%%')
    path = os.path.join(CHART_FOLDER, 'chart.png')
    plt.savefig(path)
    plt.close()

    slide.shapes.add_picture(path, Inches(2), Inches(2), height=Inches(4))
    prs.save(output_path)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
