from fastapi import FastAPI, UploadFile, File
import easyocr
import numpy as np
import cv2

app = FastAPI()

reader = easyocr.Reader(['en'])

@app.post("/extract")
async def extract_text(file: UploadFile = File(...)):
    contents = await file.read()
    nparr = np.frombuffer(contents, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    result = reader.readtext(img)

    extracted_data = []
    for (bbox, text, prob) in result:
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        center_x = sum(xs) / 4.0
        center_y = sum(ys) / 4.0

        clean_text = ''.join(filter(str.isdigit, text))

        if clean_text:
            extracted_data.append({
                "text": clean_text,
                "confidence": float(prob),
                "center_x": center_x,
                "center_y": center_y
            })

    return {"data": extracted_data}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)