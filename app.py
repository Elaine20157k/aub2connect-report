UPLOAD_FOLDER = 'uploads'

import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Circle
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from flask import Flask, request, render_template_string, send_file

# 检查 uploads 目录
if os.path.exists(UPLOAD_FOLDER) and not os.path.isdir(UPLOAD_FOLDER):
    os.remove(UPLOAD_FOLDER)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 简单上传页面模板
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

    filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filename)

    # 调用报告生成函数
    output_path = os.path.join("output", "report.pptx")
    logo_path = "logo.png"  # 如果没有可以注释掉相关部分
    generate_ppt_report(filename, output_path, logo_path)

    return send_file(output_path, as_attachment=True)

def generate_ppt_report(excel_path, output_path, logo_path):
    df = pd.read_excel(excel_path)

    total = len(df)
    checked_in = df['check-in'].str.upper().value_counts().get('Y', 0)
    not_checked_in = df['check-in'].str.upper().value_counts().get('N', 0)
    attendance_rate = round((checked_in / total) * 100, 2)

    job_counts = df['title'].value_counts()
    company_counts = df['company'].value_counts()
    no_show_df = pd.DataFrame()
    if 'check-in' in df.columns:
        no_show_df = df[df['check-in'].str.upper() == 'N']

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_slide_layout)

    # 饼图：签到情况
    fig, ax = plt.subplots()
    ax.pie([checked_in, not_checked_in], labels=['Checked In', 'Not Checked'], autopct='%1.1f%%', startangle=90)
    pie_path = os.path.join("charts", "attendance.png")
    os.makedirs("charts", exist_ok=True)
    plt.savefig(pie_path)
    plt.close()

    left = Inches(0.5)
    top = Inches(1)
    height = Inches(4)
    pic = slide.shapes.add_picture(pie_path, left, top, height=height)

    prs.save(output_path)

# 关键入口
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
