import os, json
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from agent import run_agent
from schema import ExtractedInfo

SUPPORTED = {".pdf", ".docx"}

def process_single(file_path: str, ocr_preset: str, lang: str) -> dict:
    try:
        result: ExtractedInfo = run_agent(file_path, ocr_preset=ocr_preset, lang=lang)
        return {"file": file_path, "status": "success", "data": result.model_dump()}
    except Exception as e:
        return {"file": file_path, "status": "error", "error": str(e)}

def batch_process(folder: str, max_workers: int = 4, ocr_preset: str = "scanner",
                  lang: str = "eng", output_file: str = "batch_results.json"):
    files = [str(f) for f in Path(folder).iterdir() if f.suffix.lower() in SUPPORTED]
    if not files:
        print(f"No supported files in '{folder}'")
        return
    print(f"{len(files)} file(s) found. Processing...\n")
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_single, f, ocr_preset, lang): f for f in files}
        for future in as_completed(futures):
            r = future.result()
            print(f"  [{'OK' if r['status'] == 'success' else 'FAIL'}] {Path(r['file']).name}")
            results.append(r)
    with open(output_file, "w") as f:
        json.dump(results, f, indent=2)
    success = sum(1 for r in results if r["status"] == "success")
    print(f"\nDone: {success}/{len(files)} succeeded -> '{output_file}'")

if __name__ == "__main__":
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else ".\\documents"
    preset = sys.argv[2] if len(sys.argv) > 2 else "scanner"
    lang   = sys.argv[3] if len(sys.argv) > 3 else "eng"
    batch_process(folder, ocr_preset=preset, lang=lang)