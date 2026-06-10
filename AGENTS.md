# AGENTS.md

## Collaboration

- Explain the intended changes before editing files.
- Confirm with the user before implementation when requirements, business logic,
  compatibility expectations, or behavior are unclear.
- Do not modify code unrelated to the current task.

## Project Overview

This is a Python 3.10+ document information extraction project supporting:

- PDF and DOCX documents
- Native extraction from text-based PDFs
- OCR for scanned PDFs
- A Streamlit web interface
- A FastAPI service
- Multiple AI model providers
- JSON, TXT, CSV, DOCX, and Excel exports

The project uses LangChain to call OpenAI, Anthropic, Google Gemini, DeepSeek,
xAI, and Alibaba Qwen. Extracted data is validated with Pydantic.

## Main Files

- `app_streamlit.py`: Main web application and user interactions.
- `app_fastapi.py`: FastAPI endpoints.
- `agent.py`: Model definitions, provider routing, prompts, and structured output.
- `extractor.py`: PDF, DOCX, and OCR text extraction.
- `image_enhancer.py`: Scanned-document image enhancement.
- `schema.py`: Pydantic output models.
- `batch_processor.py`: Directory batch processing.
- `translations.py`: Multilingual interface strings.
- `requirements.txt`: Python dependencies.

## Environment and Startup

Use the existing virtual environment:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Start Streamlit:

```bash
.venv/bin/streamlit run app_streamlit.py
```

The default URL is `http://localhost:8501`.

Start FastAPI:

```bash
.venv/bin/uvicorn app_fastapi:app --reload
```

API documentation is available at `http://localhost:8000/docs`.

## Environment Variables

API keys are stored in `.env`. Supported variables are:

- `OPENAI_API_KEY`
- `ANTHROPIC_API_KEY`
- `GOOGLE_API_KEY`
- `DASHSCOPE_API_KEY`
- `DEEPSEEK_API_KEY`
- `XAI_API_KEY`

Rules:

- Never display, log, commit, or expose real API keys.
- Never add `.env` to Git.
- Update `.env.example` when adding an environment variable.
- Do not call real paid models during tests unless the user explicitly requests it.

## Development Conventions

- Preserve the current single-directory Python project structure unless a
  structural change is explicitly requested.
- Reuse existing functions and dependencies instead of duplicating document
  extraction or model-routing logic.
- AI extraction results must remain compatible with `schema.ExtractedInfo`.
- When adding a model, update both `AVAILABLE_MODELS` and `PROVIDER_ENV_KEYS`.
- When adding an OCR preset, update `image_enhancer.PRESETS` and the relevant UI.
- Keep PDF and DOCX behavior consistent across Streamlit, FastAPI, and batch
  processing.
- Clean up temporary uploaded files after both successful and failed processing.
- Consider all entries in `translations.py` when changing user-facing text.
- Do not hardcode secrets, user document paths, or machine-specific settings.
- Update `requirements.txt` when dependencies change and avoid duplicate entries.

## Platform Notes

`extractor.py` currently contains Windows-specific Tesseract and Poppler paths.
OCR changes should avoid adding more hardcoded paths and should support both
Windows and Linux.

OCR system dependencies include:

- Tesseract OCR
- Poppler
- Required Tesseract language packs such as `eng` and `chi_sim`

Poppler is available in the current Linux environment, but Tesseract has not
been confirmed. Successful Python imports alone do not prove that OCR works.

## Verification

After changing Python code, run at least:

```bash
.venv/bin/python -m py_compile \
  agent.py \
  app_fastapi.py \
  app_streamlit.py \
  batch_processor.py \
  extractor.py \
  image_enhancer.py \
  schema.py \
  translations.py
```

Verify core imports:

```bash
.venv/bin/python -c \
  "import agent, app_fastapi, extractor, schema; print('imports ok')"
```

Add validation according to the affected area:

- Streamlit changes: start the application and inspect the affected UI flow.
- FastAPI changes: check `/health` and `/docs`.
- Extraction changes: test a text PDF, scanned PDF, and DOCX.
- Export changes: generate files and verify that they open successfully.
- Model-routing changes: use mocks to verify provider, model, and API-key mapping.

The project currently has no formal automated test configuration. Add focused
`pytest` tests for complex new logic and avoid relying on real external APIs.

## Git and Change Safety

- Check `git status` before editing.
- Do not overwrite or revert existing user changes.
- Before finishing, check for unresolved Git conflict markers:

```bash
rg -n '^(<<<<<<<|=======|>>>>>>>)' --glob '*.py'
```

- Do not commit `.env`, uploaded documents, generated results, caches, or logs.
- Do not commit, push, or run destructive Git commands unless explicitly asked.
