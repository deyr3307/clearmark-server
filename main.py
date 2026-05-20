import os
import fitz
from pptx import Presentation
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
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

def clean_pdf(input_path, output_path, crop_pixels=60):
    try:
        doc = fitz.open(input_path)
        for page in doc:
            rect = page.rect
            # Crop 60 pixels from the bottom of each page to remove logo and text
            new_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - crop_pixels)
            page.set_cropbox(new_rect)
        
        # Save with garbage collection to permanently remove the cropped data
        doc.save(output_path, garbage=3, deflate=True)
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
    watermark_text: str = Form(default="NotebookLM")
):
    ext = os.path.splitext(file.filename)[1].lower()
    input_path = os.path.join(UPLOAD_FOLDER, f"input_file{ext}")
    output_path = os.path.join(UPLOAD_FOLDER, f"output_file{ext}")
    
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if ext == '.pdf':
            # PDF automatically crops the bottom, no text matching needed
            clean_pdf(input_path, output_path, crop_pixels=60)
            media_type = 'application/pdf'
        elif ext in ['.pptx', '.ppt']:
            # PPTX still uses text matching
            text_to_remove = watermark_text.strip() if watermark_text.strip() else "NotebookLM"
            clean_pptx(input_path, output_path, text_to_remove)
            media_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        else:
            raise HTTPException(status_code=400, detail="Only PDF and PPTX files are supported now.")

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
    
