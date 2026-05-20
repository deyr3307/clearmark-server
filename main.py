from fastapi import FastAPI, File, UploadFile
from fastapi.responses import Response
import cv2
import numpy as np

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "ClearMark Backend is Running Perfectly!"}

@app.post("/remove-watermark/")
async def remove_watermark(file: UploadFile = File(...)):
    contents = await file.read()
    nparr = np.frombuffer(contents, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, mask = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY)
    result = cv2.inpaint(img, mask, 3, cv2.INPAINT_TELEA)

    _, encoded_img = cv2.imencode('.png', result)
    return Response(content=encoded_img.tobytes(), media_type="image/png")
  
