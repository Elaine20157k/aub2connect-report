import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Circle
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.dml.color import RGBColor
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
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

    left = Inches(0.5)
    top = Inches(0.5)
    width = Inches(2)
    height = Inches(2)
    slide.shapes.add_picture(logo_path, left, top, width=width, height=height)

    shapes = slide.shapes
    title_shape = shapes.title
    title_shape.text = "Event Dashboard Summary"

    txBox = slide.shapes.add_textbox(Inches(3), Inches(0.5), Inches(5), Inches(2))
    tf = txBox.text_frame
    tf.text = f"Total Registered: {total}\nChecked In: {checked_in}\nAttendance Rate: {attendance_rate}%\nNo-shows: {not_checked_in}"

    fig, ax = plt.subplots(figsize=(3, 3))
    labels = ['Checked In', 'Not Checked In']
    sizes = [checked_in, not_checked_in]
    colors = ['#f6a01a', '#f15946']
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)

