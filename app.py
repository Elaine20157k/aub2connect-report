import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
from pptx import Presentation
from pptx.util import Inches, Pt
from flask import Flask, request, render_template_string, send_file

# Constants
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
CHART_FOLDER = "charts"

# Ensure folders exist
for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER, CHART_FOLDER]:
    if os.path.exists(folder) and not os.path.isdir(folder):
        os.remove(folder)
    os.makedirs(folder, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# HTML for file upload page
HTML_FORM = """
<!doctype html>
<title>Upload Excel</title>
<h1>Upload Excel File to Generate PPT Report</h1>
<form method=post enctype=multipart/form-data action="/upload">
  <input type=file name=file required>
  <input type=submit value="Upload and Generate">
</form>
"""

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_FORM)

@app.route("/upload", methods=["POST"])
def upload_file():
    file = request.files.get("file")
    if not file or file.filename == '':
        return "No file selected.", 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(OUTPUT_FOLDER, "report.pptx")

    file.save(input_path)

    try:
        generate_report(input_path, output_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Error generating report: {str(e)}", 500

def generate_report(excel_path, output_path):
    df = pd.read_excel(excel_path)

    total = len(df)
    checked_in = df["check-in"].str.upper().value_counts().get("Y", 0)
    not_checked_in = df["check-in"].str.upper().value_counts().get("N", 0)
    attendance_rate = round((checked_in / total) * 100, 2)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # Create pie chart
    pie_path = os.path.join(CHART_FOLDER, "pie.png")
    fig, ax = plt.subplots()
    ax.pie([checked_in, not_checked_in],
           labels=["Checked In", "Not Checked"],
           autopct="%1.1f%%",
           startangle=90)
    centre_circle = Circle((0, 0), 0.70, fc="white")
    fig.gca().add_artist(centre_circle)
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig(pie_path)
    plt.close()

    # Insert pie chart into PPT
    slide.shapes.add_picture(pie_path, Inches(1), Inches(1.5), height=Inches(4))

    # Save PPT
    prs.save(output_path)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
