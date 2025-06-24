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

UPLOAD_FOLDER = 'uploads'
CHARTS_FOLDER = 'charts'
OUTPUT_FOLDER = 'output'

# Create necessary folders if not exist
for folder in [UPLOAD_FOLDER, CHARTS_FOLDER, OUTPUT_FOLDER]:
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

    input_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(input_path)

    output_path = os.path.join(OUTPUT_FOLDER, 'report.pptx')
    logo_path = 'logo.png'  # Optional: include if available
    generate_ppt_report(input_path, output_path, logo_path)

    return send_file(output_path, as_attachment=True)

def generate_ppt_report(excel_path, output_path, logo_path):
    df = pd.read_excel(excel_path)

    total = len(df)
    checked_in = df['check-in'].str.upper().value_counts().get('Y', 0)
    not_checked_in = df['check-in'].str.upper().value_counts().get('N', 0)
    attendance_rate = round((checked_in / total) * 100, 2) if total else 0

    job_counts = df['title'].value_counts()
    no_show_df = df[df['check-in'].str.upper() == 'N'] if 'check-in' in df.columns else pd.DataFrame()

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # Pie chart
    fig1, ax1 = plt.subplots()
    ax1.pie([checked_in, not_checked_in], labels=['Checked In', 'Not Checked'], autopct='%1.1f%%', startangle=90)
    pie_path = os.path.join(CHARTS_FOLDER, 'attendance.png')
    plt.savefig(pie_path)
    plt.close()

    # Bar chart
    fig2, ax2 = plt.subplots()
    job_counts.plot(kind='barh', ax=ax2)
    bar_path = os.path.join(CHARTS_FOLDER, 'job_distribution.png')
    plt.savefig(bar_path)
    plt.close()

    # Slide content
    slide.shapes.add_picture(pie_path, Inches(0.5), Inches(1), height=Inches(3.5))
    slide.shapes.add_picture(bar_path, Inches(4.5), Inches(1), height=Inches(3.5))

    textbox = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(9), Inches(1))
    tf = textbox.text_frame
    tf.word_wrap = True
    p = tf.add_paragraph()
    p.text = f"Total Registered: {total} | Checked In: {checked_in} | Attendance Rate: {attendance_rate}%"
    p.font.size = Pt(14)
    p.font.bold = True

    prs.save(output_path)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
