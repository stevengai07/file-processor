import tempfile, os
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from agent import run_agent
from schema import ExtractedInfo
from image_enhancer import PRESETS

app = FastAPI(title="Document Extractor API", version="2.0.0")
ALLOWED = {".pdf", ".docx"}

@app.post("/extract", response_model=ExtractedInfo)
async def extract_document(
    file: UploadFile = File(...),
    ocr_preset: str = Query("scanner", enum=list(PRESETS.keys())),
    lang: str = Query("eng")
):
    suffix = os.path.splitext(file.filename)[-1].lower()
    if suffix not in ALLOWED:
        raise HTTPException(400, f"Unsupported type '{suffix}'. Use .pdf or .docx")
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        return run_agent(tmp_path, ocr_preset=ocr_preset, lang=lang)
    except Exception as e:
        raise HTTPException(500, str(e))
    finally:
        os.unlink(tmp_path)

@app.get("/health")
def health():
    return {"status": "ok", "presets": list(PRESETS.keys())}
