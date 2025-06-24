import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.patches import Circle
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.dml.color import RGBColor

def generate_ppt_report(excel_path, output_path, logo_path):
    df = pd.read_excel(excel_path)

    total = len(df)
    checked_in = df['check-in'].str.upper().value_counts().get('Y', 0)
    not_checked_in = df['check-in'].str.upper().value_counts().get('N', 0)
    attendance_rate = round((checked_in / total) * 100, 2)
    job_counts = df['tile'].value_counts()
    company_counts = df['company'].value_counts()
    no_show_df =
        for i, row in enumerate(no_show_companies.values.tolist()):
        for j, cell in enumerate(row):
            table.cell(i + 1, j).text = str(cell)

    note_box = slide.shapes.add_textbox(Inches(8.7), Inches(5.0), Inches(4.6), Inches(0.7))
    note_frame = note_box.text_frame
    p = note_frame.add_paragraph()
    p.text = "📌 Follow-up Recommendation:\nConsider reaching out to participants who did not attend the event. You may offer a summary or share key materials."
    for para in note_frame.paragraphs:
        para.font.size = Pt(10)
        para.font.color.rgb = RGBColor(150, 75, 0)

    summary_box = slide.shapes.add_textbox(Inches(0.3), Inches(5.9), Inches(12.5), Inches(1.3))
    frame = summary_box.text_frame
    frame.text = "Summary\n"
    p = frame.add_paragraph()
    p.text = summary_text
    frame.paragraphs[0].font.size = Pt(14)
    frame.paragraphs[0].font.bold = True
    for para in frame.paragraphs[1:]:
        para.font.size = Pt(11)

    prs.save(output_path)
