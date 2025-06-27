from flask import Flask, request, jsonify
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)  # ✅ 启用 CORS 允许前端跨域访问

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def home():
    return 'Flask backend is running.'

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    print(f"✅ Received and saved: {file.filename}")
    return jsonify({'message': 'File uploaded successfully'})

@app.route('/download', methods=['GET'])
def download_report():
    # 你可以根据需求返回一个真实的 PDF 报告文件
    return jsonify({'message': 'Report download not implemented yet'})
