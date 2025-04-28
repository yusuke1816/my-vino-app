from flask import Flask, request, jsonify
from flask_cors import CORS  # CORSをインポート
import os
from pptx import Presentation

app = Flask(__name__)

# localhost:3000 からのリクエストを許可
CORS(app, resources={r"/upload": {"origins": "http://localhost:3000"}})

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if not file.filename.endswith('.pptx'):
        return jsonify({'error': 'Invalid file format. Only .pptx files are allowed'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    ppt_text = extract_text_from_ppt(file_path)

    return jsonify({'ppt_text': ppt_text})

def extract_text_from_ppt(file_path):
    try:
        prs = Presentation(file_path)
        text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text.append(shape.text)
        return "\n".join(text)
    except Exception as e:
        print(f"Error while extracting text from PPT: {e}")  # エラー内容をログに出力
        return jsonify({'error': str(e)}), 500  # エラーメッセージをクライアントに返す

if __name__ == '__main__':
    app.run(debug=True, port=5001, host='0.0.0.0')
