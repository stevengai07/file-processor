# AI-Powered File Processing System

> [中文](README.md) | English

A document-processing application built with Streamlit, FastAPI, Ollama, and Tesseract. It uses Excel templates to define fields, extracts structured information from PDF and DOCX files in batches, and provides workflows for result review, Office document export, Excel merging, and ordered document packaging.

The primary web entry point is `app_streamlit_v2.py`. All AI-assisted extraction, chat, and document-matching features currently use the Ollama model `qwen3.6:latest`. No cloud-model API key is required.

## Features

The Streamlit interface contains six workspaces:

1. **Template Configuration**: Upload an `.xlsx` template, select a worksheet and header row, and define extraction fields, field types, and required-field rules.
2. **Batch Extraction**: Process PDF and DOCX files concurrently. Text-based PDFs use their native text layer, while scanned pages automatically fall back to Tesseract OCR.
3. **Result Review**: Review successful, needs-review, and failed results; edit extracted fields manually; and export XLSX or DOCX files.
4. **AI Console**: Ask questions, create summaries, extract metrics and action items, or generate reports based on a reference document. Results can be exported as DOCX, XLSX, or PPTX.
5. **Excel Merge**: Merge multiple structurally similar `.xlsx` files, detect headers, align columns, and optionally remove duplicates. This workflow does not call the AI model.
6. **Ordered Packaging**: Match PDF, DOCX, PPT, or PPTX materials against an Excel ordering list, review uncertain matches, and generate a matching report and a ZIP archive with renamed files.

The project also provides a FastAPI service for template management, batch tasks, result review, and exports.

## File Formats

| Workflow | Input | Output |
| --- | --- | --- |
| Template configuration | XLSX | In-memory template snapshot |
| Batch extraction | PDF, DOCX | Structured results |
| Result review | Batch extraction results | XLSX, DOCX |
| AI console | Reference: PDF/DOCX; targets: PDF/DOCX | DOCX, XLSX, PPTX |
| Excel merge | Two or more XLSX files | XLSX |
| Ordered packaging | XLSX ordering list; PDF/DOCX/PPT/PPTX materials | XLSX matching report, renamed ZIP archive |

Format limitations:

- Batch extraction does not accept image files directly. Scanned documents should be provided as PDFs.
- Legacy `.doc` files have not been verified and are not considered supported.
- In ordered packaging, legacy `.ppt` files support limited filename-based matching only; `.pptx` body text can be extracted.
- The AI console interface allows a TXT reference file to be selected, but the shared extractor does not currently parse TXT files, so this format is not recommended.

## Requirements

- Python 3.10 or later
- Access to an Ollama service where `qwen3.6:latest` is available
- Tesseract OCR with its executable available on the system `PATH`
- Tesseract language packs for the selected languages, such as `eng`, `chi_sim`, `chi_tra`, `jpn`, or `kor`

Check the available Ollama models and Tesseract installation with:

```bash
ollama list
tesseract --version
tesseract --list-langs
```

OCR is optional when processing only DOCX files and PDFs with a complete text layer. Scanned PDFs require Tesseract and the appropriate language packs.

## Installation

### 1. Create a virtual environment

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 2. Install dependencies

```bash
python -m pip install -r requirements.txt
```

### 3. Configure Ollama

Create or edit `.env` in the repository root:

```dotenv
OLLAMA_BASE_URL=http://127.0.0.1:11434/v1
```

`OLLAMA_BASE_URL` must point to Ollama's OpenAI-compatible endpoint and include `/v1`. If the variable is not set, the application uses the local address shown above. The model name is currently fixed to `qwen3.6:latest`.

Do not commit `.env` files that contain internal service addresses or secrets.

## Running Streamlit

macOS / Linux:

```bash
.venv/bin/streamlit run app_streamlit_v2.py
```

Windows PowerShell:

```powershell
.venv\Scripts\streamlit.exe run app_streamlit_v2.py
```

The repository's default URL is:

```text
http://localhost:18601
```

To override the port for one run:

```bash
.venv/bin/streamlit run app_streamlit_v2.py --server.port 8501
```

## Basic Workflow

1. Open **Template Configuration**, upload an Excel template, and confirm the worksheet, header row, and field definitions.
2. Open **Batch Extraction**, upload PDF or DOCX files, select the OCR language and scanned-document preset, and start the task.
3. Open **Result Review**, inspect fields that need review, correct results, and select an export scope.
4. Download the XLSX summary or DOCX report.

The OCR presets are Scanner, Photo, Mixed, and Off. OCR is used only for PDF pages that contain insufficient native text; DOCX files are parsed directly from their document structure.

## Running FastAPI

macOS / Linux:

```bash
.venv/bin/uvicorn app_fastapi:app --reload
```

Windows PowerShell:

```powershell
.venv\Scripts\uvicorn.exe app_fastapi:app --reload
```

- Service: `http://localhost:8000`
- Health check: `http://localhost:8000/health`
- OpenAPI documentation: `http://localhost:8000/docs`

Main endpoint groups:

| Category | Endpoints |
| --- | --- |
| System | `GET /health` |
| Templates | `POST /api/templates/parse`, `POST /api/templates`, `GET /api/templates` |
| Tasks | `POST /api/tasks`, plus file upload, start, cancel, query, and retry endpoints |
| Results | Result lists, single-file details, and manual-edit endpoints |
| Exports | `POST /api/tasks/{task_id}/exports/excel`, `POST /api/tasks/{task_id}/exports/docx` |

See `/docs` for complete request parameters and response schemas.

## Merging an Excel Folder

In addition to the web interface, all `.xlsx` files in a directory can be merged from the command line:

```bash
.venv/bin/python excel_merge.py /path/to/excel-folder
```

By default, the command creates `汇总结果.xlsx` in the target directory. Command-line mode reads the first worksheet from each file, detects headers, merges columns, and adds sequential row numbers.

## Project Structure

```text
app_streamlit_v2.py   # Primary Streamlit interface
app_fastapi.py        # FastAPI service
agent.py              # Ollama calls, prompts, and structured result handling
extractor.py          # PDF, DOCX, and OCR text extraction
image_enhancer.py     # Scanned-page image enhancement
schema.py             # Pydantic data models
template_service.py   # Excel template parsing and template snapshots
task_engine.py        # Batch tasks and concurrent processing
export_service.py     # XLSX, DOCX, and PPTX exports
excel_merge.py        # Excel merge core and command-line entry point
excel_merge_ui.py     # Excel merge interface
order_matching.py     # Ordering-list parsing, document matching, and ZIP packaging
tests/                 # Automated tests
manual_test_cases/     # Manual test samples
```

`app_streamlit.py` and `batch_processor.py` are legacy entry points and are not the recommended way to run the current application.

## Current Limitations

- Templates, tasks, and results are stored in process memory and are lost when the service restarts. The current implementation is not a persistent production task system.
- The FastAPI task-start endpoint currently performs extraction synchronously, so long-running tasks may keep the request open.
- Temporary copies of uploaded files are retained to support retries. The deployment environment must manage process lifetime and temporary-directory cleanup.
- The AI console sends at most the first 6,000 characters of a reference template and the first 16,000 characters of target-document context to the model.
- OCR quality depends on scan clarity, image orientation, and installed language packs. OCR and AI output should always be reviewed by a person.
- FastAPI currently allows all CORS origins. Restrict this configuration to the actual frontend origins before a public deployment.

## Verification

Run the automated tests:

```bash
.venv/bin/python -m pytest
```

Check the syntax of the core Python files:

```bash
.venv/bin/python -m py_compile agent.py app_fastapi.py app_streamlit.py app_streamlit_v2.py batch_processor.py extractor.py image_enhancer.py schema.py translations.py
```

See `manual_test_cases/ordering_workflow/README.md` for ordered-packaging samples and expected results.
