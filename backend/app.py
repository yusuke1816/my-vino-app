from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from pptx import Presentation
from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

app = Flask(__name__)

# CORS設定（全てのオリジンを許可）
CORS(app, resources={r"/upload": {"origins": "*"}})

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

model_name = "elyza/ELYZA-japanese-Llama-2-7b-instruct"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForCausalLM.from_pretrained(model_name, torch_dtype=torch.float32, device_map={"": "cpu"})

def extract_text_from_ppt(file_path):
    prs = Presentation(file_path)
    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)
    return "\n".join(text)

def rewrite_text(text):
    prompt = f"以下の文章を丁寧な日本語に書き直してください。\n{text.strip()}\n丁寧な文章:"
    inputs = tokenizer(prompt, return_tensors="pt", truncation=True, max_length=1024)  # max_lengthを短縮
    with torch.no_grad():
        output = model.generate(
            **inputs,
            max_new_tokens=200,
            temperature=0.6,  
            do_sample=True,
            top_p=0.9,  # top_pを追加して精度向上を狙う
            top_k=50     # top_kを追加して生成の多様性を調整
        )
    result = tokenizer.decode(output[0], skip_special_tokens=True)
    
    # ここで書き直し結果をログに出力
    print("Original Text:", text)
    print("Refined Text:", result)
    
    return result.replace(prompt, "").strip()

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

    try:
        ppt_text = extract_text_from_ppt(file_path)
        refined_text = rewrite_text(ppt_text)  # ここで丁寧な日本語に変換
        return jsonify({'ppt_text': refined_text})
    except Exception as e:
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001, host='0.0.0.0', use_reloader=False)
