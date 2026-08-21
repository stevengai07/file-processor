# AI 文件智能处理系统

基于 Streamlit、FastAPI、Ollama 和 Tesseract 的本地文档处理工具。系统支持通过 Excel 模板定义字段，批量提取 PDF/DOCX 中的结构化信息，并提供结果审核、Office 文件导出、Excel 合并和发言材料顺序打包等工作流。

当前正式 Web 入口为 `app_streamlit_v2.py`，所有 AI 提取、对话和文档匹配功能固定使用 Ollama 模型 `qwen3.6:latest`，无需云端模型 API Key。

> [English summary](#english-summary)

## 主要功能

Streamlit 界面包含六个工作区：

1. **模板配置**：上传 `.xlsx` 模板，选择工作表和表头行，定义待提取字段、类型及必填规则。
2. **批量提取**：并发处理 PDF 和 DOCX 文件；文本型 PDF 使用原生文本层，扫描页自动回退到 Tesseract OCR。
3. **结果审核**：查看成功、待复核和失败结果，人工修改字段，并导出 XLSX 或 DOCX。
4. **AI 控制台**：针对多份 PDF/DOCX 进行问答、摘要、指标提取、行动项识别和参考样板仿写，可导出 DOCX、XLSX 或 PPTX。
5. **Excel 合并**：合并多个结构相近的 `.xlsx` 文件，自动检测表头、对齐字段并可选去重；该功能不调用 AI。
6. **顺序编号**：根据发言顺序 Excel 匹配 PDF、DOCX、PPT/PPTX 材料，人工复核后生成匹配清单和重命名 ZIP。

项目同时提供 FastAPI 接口，用于模板管理、批量任务、结果审核和导出。

## 文件格式

| 工作流 | 输入 | 输出 |
| --- | --- | --- |
| 模板配置 | XLSX | 内存中的模板快照 |
| 批量提取 | PDF、DOCX | 结构化结果 |
| 结果审核 | 批量提取结果 | XLSX、DOCX |
| AI 控制台 | 参考样板：PDF/DOCX；目标文档：PDF/DOCX | DOCX、XLSX、PPTX |
| Excel 合并 | 两个或以上 XLSX | XLSX |
| 顺序编号 | 顺序表 XLSX；材料 PDF/DOCX/PPT/PPTX | 匹配清单 XLSX、重命名 ZIP |

限制说明：

- 批量提取不支持直接上传图片，扫描件应封装为 PDF。
- 旧版 `.doc` 未验证，不作为支持格式。
- 顺序编号中的旧版 `.ppt` 只能进行有限的文件名匹配；`.pptx` 可提取正文。
- AI 控制台界面允许选择 TXT 参考样板，但当前通用提取器未实现 TXT 解析，因此不建议使用。

## 环境要求

- Python 3.10 或更高版本
- 可访问的 Ollama 服务，且服务中已提供 `qwen3.6:latest`
- Tesseract OCR，可执行文件需位于系统 `PATH`
- 与所选语言对应的 Tesseract 语言包，例如 `eng`、`chi_sim`、`chi_tra`、`jpn` 或 `kor`

可通过以下命令检查 Ollama 模型和 Tesseract：

```bash
ollama list
tesseract --version
tesseract --list-langs
```

如果仅处理带有完整文本层的 PDF 和 DOCX，可以不使用 OCR；扫描 PDF 需要正确安装 Tesseract 及语言包。

## 安装

### 1. 创建虚拟环境

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell：

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 2. 安装依赖

```bash
python -m pip install -r requirements.txt
```

### 3. 配置 Ollama

在项目根目录创建或修改 `.env`：

```dotenv
OLLAMA_BASE_URL=http://127.0.0.1:11434/v1
```

`OLLAMA_BASE_URL` 必须是 Ollama 的 OpenAI 兼容接口地址，并包含 `/v1`。未设置该变量时，程序使用上面的本机地址。模型名称当前固定为 `qwen3.6:latest`。

不要提交包含内部服务地址或密钥的 `.env` 文件。

## 启动 Streamlit

macOS / Linux：

```bash
.venv/bin/streamlit run app_streamlit_v2.py
```

Windows PowerShell：

```powershell
.venv\Scripts\streamlit.exe run app_streamlit_v2.py
```

仓库配置的默认访问地址为：

```text
http://localhost:18601
```

如需临时覆盖端口：

```bash
.venv/bin/streamlit run app_streamlit_v2.py --server.port 8501
```

## 基本使用流程

1. 在“模板配置”上传 Excel 模板，确认工作表、表头和字段定义。
2. 在“批量提取”上传 PDF/DOCX，选择 OCR 语言和扫描文件预设后启动任务。
3. 在“结果审核”检查待复核字段、修正结果并选择导出范围。
4. 下载 XLSX 汇总表或 DOCX 报告。

OCR 预设包括扫描仪、照片、混合和关闭。OCR 只用于原生文本不足的 PDF 页面；DOCX 使用文档结构直接解析。

## 启动 FastAPI

macOS / Linux：

```bash
.venv/bin/uvicorn app_fastapi:app --reload
```

Windows PowerShell：

```powershell
.venv\Scripts\uvicorn.exe app_fastapi:app --reload
```

- 服务地址：`http://localhost:8000`
- 健康检查：`http://localhost:8000/health`
- OpenAPI 文档：`http://localhost:8000/docs`

主要接口：

| 类别 | 接口 |
| --- | --- |
| 系统 | `GET /health` |
| 模板 | `POST /api/templates/parse`、`POST /api/templates`、`GET /api/templates` |
| 任务 | `POST /api/tasks`、文件上传、启动、取消、查询和重试接口 |
| 结果 | 结果列表、单文件详情和人工修改接口 |
| 导出 | `POST /api/tasks/{task_id}/exports/excel`、`POST /api/tasks/{task_id}/exports/docx` |

完整请求参数和响应结构以 `/docs` 为准。

## Excel 文件夹合并

除 Web 界面外，也可以直接合并目录下的 `.xlsx` 文件：

```bash
.venv/bin/python excel_merge.py /path/to/excel-folder
```

默认在目标目录生成 `汇总结果.xlsx`。命令行模式读取每个文件的第一个工作表，自动检测表头、合并字段并添加连续序号。

## 项目结构

```text
app_streamlit_v2.py   # 当前 Streamlit 主界面
app_fastapi.py        # FastAPI 服务
agent.py              # Ollama 调用、提示词和结构化结果处理
extractor.py          # PDF、DOCX 和 OCR 文本提取
image_enhancer.py     # 扫描页图像增强
schema.py             # Pydantic 数据模型
template_service.py   # Excel 模板解析与模板快照
task_engine.py        # 批量任务和并发处理
export_service.py     # XLSX、DOCX 和 PPTX 导出
excel_merge.py        # Excel 合并核心及命令行入口
excel_merge_ui.py     # Excel 合并界面
order_matching.py     # 顺序表解析、材料匹配和 ZIP 打包
tests/                 # 自动化测试
manual_test_cases/     # 手工测试样例
```

`app_streamlit.py` 和 `batch_processor.py` 属于旧版入口，不是当前推荐运行方式。

## 当前限制

- 模板、任务和结果保存在进程内存中，服务重启后不会保留；当前实现不适合作为持久化生产任务系统。
- FastAPI 的任务启动接口当前会同步执行提取，长任务可能持续占用请求。
- 上传文件的临时副本会保留以支持失败重试，需由部署环境管理进程生命周期和临时目录清理。
- AI 控制台向模型注入参考样板时最多使用前 6,000 个字符，目标文档上下文最多使用前 16,000 个字符。
- OCR 质量取决于扫描清晰度、图像方向和已安装的语言包；OCR 或 AI 结果均应经过人工复核。
- FastAPI 当前允许所有 CORS 来源，公开部署前应按实际前端来源收紧配置。

## 验证

运行自动化测试：

```bash
.venv/bin/python -m pytest
```

检查核心 Python 文件语法：

```bash
.venv/bin/python -m py_compile agent.py app_fastapi.py app_streamlit.py app_streamlit_v2.py batch_processor.py extractor.py image_enhancer.py schema.py translations.py
```

顺序编号工作流的样例和预期结果见 `manual_test_cases/ordering_workflow/README.md`。

## English Summary

This repository provides a local document-processing application built with Streamlit, FastAPI, Ollama, and Tesseract. The primary UI entry point is `app_streamlit_v2.py`. All AI-assisted workflows currently use the Ollama model `qwen3.6:latest`; cloud-provider API keys are not required.

### Core Features

- Excel-template-driven extraction from PDF and DOCX files
- Native PDF text extraction with Tesseract OCR fallback for scanned pages
- Concurrent batch processing, validation, human review, and XLSX/DOCX export
- Multi-document AI console with DOCX, XLSX, and PPTX output
- Local Excel merging without LLM calls
- Excel-based document ordering, review, renaming, and ZIP packaging
- FastAPI endpoints for templates, tasks, results, retries, and exports

### Quick Start

Requirements: Python 3.10+, an accessible Ollama server containing `qwen3.6:latest`, and Tesseract with the required OCR language packs.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Create `.env` in the repository root:

```dotenv
OLLAMA_BASE_URL=http://127.0.0.1:11434/v1
```

Start the Streamlit application:

```bash
.venv/bin/streamlit run app_streamlit_v2.py
```

Open `http://localhost:18601`.

Start the API:

```bash
.venv/bin/uvicorn app_fastapi:app --reload
```

Open `http://localhost:8000/docs` for the API reference.

### Important Limitations

- Templates, tasks, and results are stored in process memory and are lost after restart.
- The API task-start operation currently runs synchronously.
- OCR requires the Tesseract executable and matching language packs.
- The AI console truncates template context to 6,000 characters and target-document context to 16,000 characters.
- AI and OCR output should be reviewed before operational use.
