import os
import fitz
from pptx import Presentation
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from PIL import Image, ImageDraw

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = '/tmp' if os.name != 'nt' else './uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def clean_pdf(input_path, output_path, target_text):
    try:
        doc = fitz.open(input_path)
        for page in doc:
            text_instances = page.search_for(target_text)
            for inst in text_instances:
                page.add_redact_annot(inst, fill=(1, 1, 1))
            page.apply_redactions()
        doc.save(output_path)
        doc.close()
    except Exception as e:
        raise Exception(f"PDF Processing Error: {str(e)}")

def clean_pptx(input_path, output_path, target_text):
    try:
        prs = Presentation(input_path)
        for slide in prs.slides:
            for shape in list(slide.shapes):
                if shape.has_text_frame:
                    text = shape.text.lower()
                    if target_text.lower() in text:
                        sp = shape._element
                        sp.getparent().remove(sp)
        prs.save(output_path)
    except Exception as e:
        raise Exception(f"PPTX Processing Error: {str(e)}")

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded."}), 400
        
    file = request.files['file']
    target_text = request.form.get('watermark_text', '').strip()
    
    if file.filename == '':
        return jsonify({"error": "No selected file."}), 400
        
    if not target_text:
        return jsonify({"error": "Watermark text is required."}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    input_path = os.path.join(UPLOAD_FOLDER, f"input_file{ext}")
    output_path = os.path.join(UPLOAD_FOLDER, f"output_file{ext}")
    
    file.save(input_path)

    try:
        if ext == '.pdf':
            clean_pdf(input_path, output_path, target_text)
            mimetype = 'application/pdf'
        elif ext in ['.pptx', '.ppt']:
            clean_pptx(input_path, output_path, target_text)
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        elif ext in ['.png', '.jpg', '.jpeg']:
            img = Image.open(input_path)
            draw = ImageDraw.Draw(img)
            w, h = img.size
            draw.rectangle([(0, h - 60), (w, h)], fill="white")
            img.save(output_path)
            mimetype = f"image/{ext[1:]}"
        else:
            return jsonify({"error": "Unsupported file format."}), 400

        return send_file(output_path, mimetype=mimetype, as_attachment=True, download_name=f"clean_{file.filename}")

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
        
