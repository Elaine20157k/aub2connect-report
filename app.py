from flask import Flask, request, render_template_string, send_from_directory
import os
from generate_report_aub2connect import generate_ppt_report

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
LOGO_PATH = 'logo.png'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head><title>AuB2Connect Report</title></head>
<body style="font-family: Arial; text-align: center; margin-top: 50px;">
    <h1>📊 AuB2Connect Report Generator</h1>
    <p>Upload your Excel file to generate a professional PPT report.</p>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="excel" accept=".xlsx" required><br><br>
        <input type="submit" value="Generate Report">
    </form>
    {% if download_link %}
        <p style="margin-top: 20px;">
            ✅ Report generated: <a href="{{ download_link }}" target="_blank">Download PPT</a>
        </p>
    {% endif %}
</body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files['excel']
    if file and file.filename.endswith('.xlsx'):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        output_name = f"report_{file.filename.replace('.xlsx', '')}.pptx"
        output_path = os.path.join(OUTPUT_FOLDER, output_name)

        generate_ppt_report(filepath, output_path, LOGO_PATH)
        download_url = f"/download/{output_name}"
        return render_template_string(HTML_TEMPLATE, download_link=download_url)
    return "Invalid file format. Please upload a .xlsx file.", 400

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

# ✅ 最重要的一行：为 Render 绑定端口
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
