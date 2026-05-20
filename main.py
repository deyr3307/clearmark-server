import os
import fitz
from pptx import Presentation
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from PIL import Image, ImageDraw
import shutil
import uvicorn

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    watermark_text: str = Form(...)
):
    if not watermark_text.strip():
        raise HTTPException(status_code=400, detail="Watermark text is required.")
        
    ext = os.path.splitext(file.filename)[1].lower()
    input_path = os.path.join(UPLOAD_FOLDER, f"input_file{ext}")
    output_path = os.path.join(UPLOAD_FOLDER, f"output_file{ext}")
    
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if ext == '.pdf':
            clean_pdf(input_path, output_path, watermark_text.strip())
            media_type = 'application/pdf'
        elif ext in ['.pptx', '.ppt']:
            clean_pptx(input_path, output_path, watermark_text.strip())
            media_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        elif ext in ['.png', '.jpg', '.jpeg']:
            img = Image.open(input_path)
            draw = ImageDraw.Draw(img)
            w, h = img.size
            draw.rectangle([(0, h - 60), (w, h)], fill="white")
            img.save(output_path)
            media_type = f"image/{ext[1:]}"
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format.")

        return FileResponse(
            path=output_path, 
            filename=f"clean_{file.filename}", 
            media_type=media_type
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
    
