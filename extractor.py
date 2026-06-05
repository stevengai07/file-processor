import numpy as np
import pytesseract
from pathlib import Path
from pdf2image import convert_from_path
from pypdf import PdfReader
from image_enhancer import enhance_document_image, PRESETS, EnhancerConfig

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

MIN_TEXT_CHARS = 100

def extract_text(file_path: str, ocr_preset: str = "scanner", lang: str = "eng") -> str:
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return _extract_pdf_smart(file_path, ocr_preset, lang)
    elif ext == ".docx":
        return _extract_docx(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .pdf or .docx")

def _extract_pdf_smart(file_path, ocr_preset, lang):
    native = _extract_pdf_native(file_path)
    char_count = len(native.replace(" ", "").replace("\n", ""))
    if char_count >= MIN_TEXT_CHARS:
        print(f"Text-based PDF ({char_count} chars extracted natively).")
        return native
    print(f"Scanned PDF detected. Running OCR [{ocr_preset}]...")
    return _extract_pdf_ocr(file_path, ocr_preset, lang)

def _extract_pdf_native(file_path):
    reader = PdfReader(file_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages).strip()

def _extract_pdf_ocr(file_path, ocr_preset="scanner", lang="eng", dpi=300):
    config = PRESETS.get(ocr_preset, EnhancerConfig())
    pages = convert_from_path(
        file_path,
        dpi=dpi,
        poppler_path=r"C:\poppler\poppler-26.02.0\Library\bin"
    )
    print(f"{len(pages)} page(s) found.")
    results = []
    for i, page_img in enumerate(pages):
        print(f"OCR page {i+1}/{len(pages)}...")
        img = np.array(page_img)
        enhanced = enhance_document_image(img, config=config)
        text = pytesseract.image_to_string(enhanced, lang=lang, config="--psm 3")
        results.append(text.strip())
    return "\n\n".join(results)

def _extract_docx(file_path):
    from docx import Document
    doc = Document(file_path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())