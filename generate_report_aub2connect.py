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
