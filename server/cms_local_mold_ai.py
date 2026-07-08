from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import uuid
import json
import csv
import os
import cv2

from ultralytics import YOLO


BASE_DIR = Path(r"C:\CMS_AI")
UPLOAD_DIR = BASE_DIR / "uploads"
RESULT_DIR = BASE_DIR / "results"
MODEL_PATH = BASE_DIR / "models" / "cms_mold_yolo26.pt"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
RESULT_DIR.mkdir(parents=True, exist_ok=True)

# Use your custom mold model if it exists. Otherwise use YOLO26s.
if MODEL_PATH.exists():
    model = YOLO(str(MODEL_PATH))
    MODEL_USED = str(MODEL_PATH)
else:
    model = YOLO("yolo26s.pt")
    MODEL_USED = "yolo26s.pt fallback generic model"

app = FastAPI(title="CMS Local Mold AI")
app.mount("/results", StaticFiles(directory=str(RESULT_DIR)), name="results")


@app.get("/", response_class=HTMLResponse)
def home():
    return f"""
    <html>
    <head>
        <title>CMS Local Mold AI</title>
        <style>
            body {{ font-family: Arial; margin: 40px; }}
            .box {{ border: 1px solid #ccc; padding: 20px; width: 650px; }}
            .warn {{ color: #b00000; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h1>CMS Local Mold AI</h1>
        <p><b>Model:</b> {MODEL_USED}</p>
        <p class="warn">Local only: this upload goes to C:\\CMS_AI on this computer.</p>

        <div class="box">
            <form action="/upload" enctype="multipart/form-data" method="post">
                <p>Select mold/base JPEG or PNG:</p>
                <input name="file" type="file" accept=".jpg,.jpeg,.png">
                <br><br>
                <input type="submit" value="Analyze Image">
            </form>
        </div>
    </body>
    </html>
    """


@app.get("/health")
def health():
    return {
        "ok": True,
        "model": MODEL_USED,
        "upload_dir": str(UPLOAD_DIR),
        "result_dir": str(RESULT_DIR),
    }


@app.post("/upload", response_class=HTMLResponse)
async def upload_image(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()

    if ext not in [".jpg", ".jpeg", ".png"]:
        return HTMLResponse("<h2>Only JPG, JPEG, PNG allowed for this first version.</h2>", status_code=400)

    job_id = uuid.uuid4().hex[:12]
    input_path = UPLOAD_DIR / f"{job_id}{ext}"

    data = await file.read()
    input_path.write_bytes(data)

    result = analyze_image(input_path, job_id)

    annotated_url = f"/results/{result['annotated_name']}"
    csv_url = f"/results/{result['csv_name']}"
    json_url = f"/results/{result['json_name']}"

    rows_html = ""
    for d in result["detections"]:
        rows_html += f"""
        <tr>
            <td>{d['class_name']}</td>
            <td>{d['confidence']:.3f}</td>
            <td>{d['x1']:.1f}, {d['y1']:.1f}, {d['x2']:.1f}, {d['y2']:.1f}</td>
        </tr>
        """

    return f"""
    <html>
    <head>
        <title>CMS Local Mold AI Result</title>
        <style>
            body {{ font-family: Arial; margin: 40px; }}
            img {{ max-width: 1100px; border: 1px solid #ccc; }}
            table {{ border-collapse: collapse; margin-top: 20px; }}
            td, th {{ border: 1px solid #ccc; padding: 6px 10px; }}
        </style>
    </head>
    <body>
        <h1>Analysis Complete</h1>
        <p><b>Original file:</b> {file.filename}</p>
        <p><b>Detections:</b> {len(result['detections'])}</p>

        <p>
            <a href="{annotated_url}" target="_blank">Open annotated image</a> |
            <a href="{csv_url}" target="_blank">Download CSV</a> |
            <a href="{json_url}" target="_blank">Download JSON</a>
        </p>

        <img src="{annotated_url}">

        <h2>Detections</h2>
        <table>
            <tr>
                <th>Class</th>
                <th>Confidence</th>
                <th>Box x1,y1,x2,y2</th>
            </tr>
            {rows_html}
        </table>

        <p><a href="/">Analyze another image</a></p>
    </body>
    </html>
    """


@app.post("/predict-json")
async def predict_json(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()

    if ext not in [".jpg", ".jpeg", ".png"]:
        return JSONResponse({"ok": False, "error": "Only JPG, JPEG, PNG allowed."}, status_code=400)

    job_id = uuid.uuid4().hex[:12]
    input_path = UPLOAD_DIR / f"{job_id}{ext}"
    input_path.write_bytes(await file.read())

    result = analyze_image(input_path, job_id)
    return result


def analyze_image(input_path: Path, job_id: str):
    results = model.predict(
        source=str(input_path),
        imgsz=1024,
        conf=0.20,
        save=False,
        verbose=False,
    )

    detections = []
    annotated = None

    for result in results:
        names = result.names

        if result.boxes is not None:
            for box in result.boxes:
                cls_id = int(box.cls[0])
                conf = float(box.conf[0])
                x1, y1, x2, y2 = box.xyxy[0].cpu().numpy().tolist()

                detections.append({
                    "class_id": cls_id,
                    "class_name": names.get(cls_id, str(cls_id)),
                    "confidence": conf,
                    "x1": float(x1),
                    "y1": float(y1),
                    "x2": float(x2),
                    "y2": float(y2),
                    "width": float(x2 - x1),
                    "height": float(y2 - y1),
                })

        annotated = result.plot()

    annotated_name = f"{job_id}_YOLO_ANALYZED.jpg"
    json_name = f"{job_id}_YOLO_Detections.json"
    csv_name = f"{job_id}_YOLO_Detections.csv"

    annotated_path = RESULT_DIR / annotated_name
    json_path = RESULT_DIR / json_name
    csv_path = RESULT_DIR / csv_name

    if annotated is not None:
        cv2.imwrite(str(annotated_path), annotated)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(detections, f, indent=2)

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        fields = [
            "class_id", "class_name", "confidence",
            "x1", "y1", "x2", "y2", "width", "height"
        ]
        wr = csv.DictWriter(f, fieldnames=fields)
        wr.writeheader()
        for d in detections:
            wr.writerow(d)

    return {
        "ok": True,
        "model": MODEL_USED,
        "input_image": str(input_path),
        "annotated_image": str(annotated_path),
        "annotated_name": annotated_name,
        "json": str(json_path),
        "json_name": json_name,
        "csv": str(csv_path),
        "csv_name": csv_name,
        "detections": detections,
    }