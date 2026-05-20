import os
import fitz
from pptx import Presentation
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import uvicorn
import uuid

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


def clean_pdf(input_path, output_path, crop_pixels=90):
    try:
        doc = fitz.open(input_path)
        for page in doc:
            rect = page.rect
            new_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 - crop_pixels)
            page.set_cropbox(new_rect)
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

    uid = uuid.uuid4().hex
    input_path = os.path.join(UPLOAD_FOLDER, f"input_{uid}{ext}")
    output_path = os.path.join(UPLOAD_FOLDER, f"output_{uid}{ext}")

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        if ext == '.pdf':
            clean_pdf(input_path, output_path, crop_pixels=90)
            media_type = 'application/pdf'
        elif ext in ['.pptx', '.ppt']:
            text_to_remove = watermark_text.strip() or "NotebookLM"
            clean_pptx(input_path, output_path, text_to_remove)
            media_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        else:
            raise HTTPException(status_code=400, detail="Only PDF and PPTX files are supported.")

        return FileResponse(
            path=output_path,
            filename=f"clean_{file.filename}",
            media_type=media_type
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
