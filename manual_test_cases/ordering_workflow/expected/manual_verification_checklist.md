# Manual Verification Checklist

## Setup

- [ ] Start Streamlit with `.venv/bin/streamlit run app_streamlit_v2.py`.
- [ ] Open the `⑥ 顺序编号` page.
- [ ] Upload `manual_test_cases/ordering_workflow/ordering_sample.xlsx`.

## Excel Parsing

- [ ] Confirm the selected worksheet is `发言顺序`.
- [ ] Confirm the header row is detected as row 1.
- [ ] Confirm column mapping shows order and title columns.
- [ ] Confirm the code/material-number column is blank or `未识别`.
- [ ] Confirm the preview shows 5 valid order items.
- [ ] Confirm no Excel parsing errors are shown.

## Document Upload

- [ ] Upload every file in `manual_test_cases/ordering_workflow/docs/`.
- [ ] Confirm file count is 6.
- [ ] Select OCR language `chi_sim+eng`.
- [ ] Select OCR enhancement `scanner`.
- [ ] For deterministic validation, turn off `使用 Ollama 提取文档标题/编号` first.

## Matching Run

- [ ] Click `开始提取并匹配`.
- [ ] Confirm progress reaches completion.
- [ ] Confirm KPI `顺序项` is 5.
- [ ] Confirm KPI `输入文件` is 6.
- [ ] Confirm at least orders 001, 002, and 003 are matched or reviewable with correct title-based candidates.
- [ ] Confirm orders 004 and 005 are unmatched or require review.
- [ ] Confirm the unrelated file is unused.
- [ ] Confirm the corrupt PDF failure does not stop processing other files.

## Manual Review

- [ ] If order 001 enters review, choose `人工智能产业发展报告.pdf` and click `确认选择`.
- [ ] For order 004, click `跳过该项`.
- [ ] For order 005, click `跳过该项`.
- [ ] Confirm no unresolved `needs_review` or `unmatched` rows remain.

## Export

- [ ] Before resolving all review items, confirm `发言材料.zip` is disabled.
- [ ] After resolving/skipping all items, confirm `发言材料.zip` is enabled.
- [ ] Download `匹配清单.xlsx`.
- [ ] Download `发言材料.zip`.
- [ ] Open `匹配清单.xlsx` and confirm status, original filename, new filename, score, and method columns are present.
- [ ] Extract the ZIP and confirm ordered files start with `001_`, `002_`, and `003_`.
- [ ] Confirm source files under `docs/` were not renamed or modified.

## Optional Ollama Pass

- [ ] If local Ollama is running, repeat the workflow with `使用 Ollama 提取文档标题/编号` enabled.
- [ ] Confirm the UI handles model extraction without exposing API keys or blocking non-failed files.
