import os
import pandas as pd
import matplotlib.pyplot as plt
from flask import Flask, request, render_template_string, send_file
from pptx import Presentation
from pptx.util import Inches
import seaborn as sns

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
    logo_path = 'logo.png' if os.path.exists('logo.png') else None
    generate_ppt_report(filepath, output_path, logo_path)
    return send_file(output_path, as_attachment=True)

def generate_ppt_report(excel_path, output_path, logo_path=None):
    df = pd.read_excel(excel_path)

    total = len(df)
    checked_in = df['check-in'].str.upper().value_counts().get('Y', 0)
    not_checked_in = df['check-in'].str.upper().value_counts().get('N', 0)
    attendance_rate = round((checked_in / total) * 100, 2)

    job_counts = df['title'].value_counts()
    company_counts = df['company'].value_counts()
    no_show_df = df[df['check-in'].str.upper() == 'N']

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # 饼图
    fig, ax = plt.subplots()
    ax.pie([checked_in, not_checked_in], labels=['Checked In', 'Not Checked'], autopct='%1.1f%%', startangle=90)
    pie_path = os.path.join(CHART_FOLDER, 'attendance.png')
    plt.savefig(pie_path)
    plt.close()
    slide.shapes.add_picture(pie_path, Inches(1), Inches(1.5), height=Inches(3))

    # 柱状图
    fig, ax = plt.subplots()
    job_counts.plot(kind='barh', ax=ax)
    bar_path = os.path.join(CHART_FOLDER, 'job_title.png')
    plt.tight_layout()
    plt.savefig(bar_path)
    plt.close()
    slide.shapes.add_picture(bar_path, Inches(4.5), Inches(1.5), height=Inches(3))

    # Logo
    if logo_path:
        slide.shapes.add_picture(logo_path, Inches(0.2), Inches(0.2), height=Inches(0.6))

    # 总结页
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = f"A total of {total} participants registered. {checked_in} checked in. Attendance rate: {attendance_rate}%."

    prs.save(output_path)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
