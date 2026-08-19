# Streamlit Ordering Workflow Manual Test Cases

This directory contains small sample files for manually verifying task `6.6`: Streamlit ordering workflow with a sample Excel and PDF/DOCX set.

## Files

- `ordering_sample.xlsx`: speaking order table with 5 order items.
- `docs/人工智能产业发展报告.pdf`: expected title match for order 001.
- `docs/数字政府建设汇报.docx`: expected title match for order 002.
- `docs/绿色能源项目材料.pdf`: expected title match for order 003.
- `docs/重复候选_人工智能产业发展报告.docx`: duplicate title candidate for manual review or conflict visibility.
- `docs/unrelated_无关材料.pdf`: unused file candidate.
- `docs/unreadable_corrupt.pdf`: intentionally invalid PDF for failure isolation.

The sample intentionally does not include material codes in the Excel, filenames, or document body. This matches the common real workflow where ordering depends on title/document-name similarity rather than stable material numbers.

## Regenerate Samples

Run from the repository root:

```bash
.venv/bin/python manual_test_cases/ordering_workflow/generate_samples.py
```

The generator does not call external APIs or paid models.

## Recommended Streamlit Settings

- Page: `⑥ 顺序编号`
- OCR language: `chi_sim+eng`
- OCR enhancement: `scanner`
- Ollama extraction: off for deterministic filename/title matching, on for full local-model verification if Ollama is running

## Expected Output Artifacts

- `匹配清单.xlsx`
- `发言材料.zip`

ZIP generation should remain disabled while unresolved `needs_review`, `unmatched`, or `failed` items are present. Confirm or skip those items before downloading the ZIP.
