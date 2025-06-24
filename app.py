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

# 创建 uploads 文件夹（如果是文件则先删除）
if os.path.exists(UPLOAD_FOLDER) and not os.path.isdir(UPLOAD_FOLDER):
    os.remove(UPLOAD_FOLDER)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 上传界面 HTML 模板
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
    file = request.files.get('file')
    if not file or file.filename == '':
        return 'No file selected.'

    filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filename)

    output_path = os.path.join("output", "report.pptx")
    logo_path = "logo.png"  # 可替换成你上传的 logo 图片
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
    no_show_df = df[df['check-in'].str.upper() == 'N'] if 'check-in' in df.columns else pd.DataFrame()

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # 保存饼图
    os.makedirs("charts", exist_ok=True)
    pie_path = os.path.join("charts", "attendance.png")
    fig, ax = plt.subplots()
    ax.pie([checked_in, not_checked_in], labels=['Checked In', 'Not Checked'], autopct='%1.1f%%', startangle=90)
    plt.savefig(pie_path)
    plt.close()

    slide.shapes.add_picture(pie_path, Inches(0.5), Inches(1), height=Inches(4))

    prs.save(output_path)

# 启动服务器
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
