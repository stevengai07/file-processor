# AI Document Information Extractor

A multi-provider AI agent that extracts structured information from PDF and DOCX
files — including scanned documents — via a Streamlit web interface.

## Features
- Multi-provider AI support — OpenAI, Anthropic, Google Gemini, Alibaba Qwen, DeepSeek, xAI Grok
- Native text extraction — fast parsing of text-based PDFs and DOCX files
- Scanned document support — image-based PDFs processed via vision models
- Image enhancement — presets for scanner, phone camera, faded, and fax
- Structured JSON output via schema validation
- Multilingual UI — English, Chinese, French, Spanish, Japanese, Korean

## Project Structure
DocAgent/
├── app_streamlit.py       # Main Streamlit application
├── agent.py               # AI provider routing and model definitions
├── schema.py              # Output schema definitions
├── image_enhancer.py      # Image preprocessing presets
├── translations.py        # UI string translations
├── requirements.txt       # Python dependencies
├── .env                   # API keys (create from .env.example)
└── .env.example           # Template for environment variables

## Supported Models
| Provider  | Example Models                    | Env Key             |
|-----------|-----------------------------------|---------------------|
| OpenAI    | gpt-4o, gpt-4o-mini, gpt-4-turbo | OPENAI_API_KEY      |
| Anthropic | claude-3-5-sonnet, claude-3-haiku | ANTHROPIC_API_KEY   |
| Google    | gemini-1.5-pro, gemini-1.5-flash  | GOOGLE_API_KEY      |
| Alibaba   | qwen-turbo, qwen-vl-plus          | DASHSCOPE_API_KEY   |
| DeepSeek  | deepseek-chat, deepseek-reasoner  | DEEPSEEK_API_KEY    |
| xAI       | grok-2-vision, grok-2             | XAI_API_KEY         |

## Setup

### 1. Create and activate virtual environment

Windows (PowerShell):
  python -m venv .venv
  .venv\Scripts\Activate.ps1

macOS / Linux:
  python -m venv .venv
  source .venv/bin/activate

### 2. Install dependencies
  pip install -r requirements.txt

### 3. Configure API keys
  copy .env.example .env
  # Then edit .env and fill in your keys

### 4. Run the app
  streamlit run app_streamlit.py
  # Opens at http://localhost:8501

## Image Enhancement Presets
| Preset       | Best For                            |
|--------------|-------------------------------------|
| scanner      | High-contrast office scanner output |
| phone_camera | Photos taken with a mobile device   |
| faded        | Old or low-contrast documents       |
| fax          | Fax machine output / noisy docs     |

## Troubleshooting

Sidebar not visible:
  Remove `header { visibility: hidden; }` from the st.markdown() CSS block
  in app_streamlit.py. The toggle arrow is on the far left browser edge.

ModuleNotFoundError on startup:
  Ensure (.venv) appears in your prompt, then re-run:
  pip install -r requirements.txt

SyntaxError in PowerShell one-liners:
  Write the script to a .py file first, then run: python fix.py

API key not detected:
  Keys must be in .env in the same folder as app_streamlit.py

## Requirements
- Python 3.10 or higher
- Virtual environment (strongly recommended)
- At least one valid AI provider API key

## License
MIT