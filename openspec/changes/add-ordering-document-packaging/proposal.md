## Why

Users need to turn a meeting or speaking-order Excel sheet plus a batch of PDF/DOCX materials into an ordered, safely renamed ZIP without manually comparing every document. The current app can extract document content, but it does not support one-to-one order matching, manual review, or packaging for ordered speaking materials.

## What Changes

- Add a sixth Streamlit workflow page in `app_streamlit_v2.py` for uploading a speaking-order Excel file and PDF/DOCX documents while preserving the existing `⑤ Excel 合并` page.
- Parse order rows from `.xlsx` files with sheet/header-row selection, automatic column mapping, manual mapping fallback, validation, and original row references.
- Extract document identity from filenames, extracted text, OCR output, and Ollama-assisted structured extraction.
- Match order items to documents by material code first, title second, and Ollama only for limited ambiguous semantic decisions.
- Enforce global one-to-one assignment so one order item maps to at most one file and one file is assigned to at most one order item.
- Route low-confidence, unmatched, duplicate, or ambiguous results into manual review before export.
- Export safely renamed PDF/DOCX files plus `匹配清单.xlsx` in a ZIP without modifying original uploads.
- Keep the first implementation Streamlit-first; FastAPI ordering endpoints are deferred unless a separate follow-up change requests them.
- Fix existing blockers in extraction settings, OCR disabling, OCR failure handling, task temporary-file lifecycle, and model-used consistency.
- Add focused pytest coverage using mocked Ollama calls and in-memory files.

## Capabilities

### New Capabilities

- `ordering-document-packaging`: Parse speaking-order Excel files, extract document identities, match documents to order items, support manual review, and export ordered renamed ZIP packages with a matching report.

### Modified Capabilities

- None.

## Impact

- Affected modules: `schema.py`, `agent.py`, `extractor.py`, `task_engine.py`, `app_streamlit_v2.py`.
- New module: `order_matching.py`.
- New tests for Excel parsing, matching, manual review, ZIP/report output, Ollama failure handling, OCR failure handling, and task temporary-file lifecycle.
- Uses existing dependencies where possible: `openpyxl`, `pdfplumber`, DOCX extraction, Tesseract OCR pipeline, LangChain OpenAI-compatible Ollama calls.
- Does not send documents to cloud LLMs; Ollama URL remains environment-driven via `OLLAMA_BASE_URL`.
