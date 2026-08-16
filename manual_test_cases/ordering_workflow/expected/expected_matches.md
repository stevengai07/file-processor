# Expected Ordering Matches

| Order | Title | Expected file | Expected status | Notes |
|---:|---|---|---|---|
| 001 | 人工智能产业发展报告 | `人工智能产业发展报告.pdf` | `matched` or `needs_review` | Duplicate title candidate may require manual review depending on extracted identity. |
| 002 | 数字政府建设汇报 | `数字政府建设汇报.docx` | `matched` | Title match. |
| 003 | 绿色能源项目材料 | `绿色能源项目材料.pdf` | `matched` | Title match. |
| 004 | 智慧交通建设方案 | none | `unmatched` | Manually skip this item before ZIP export. |
| 005 | 未提交材料测试 | none | `unmatched` | Manually skip this item before ZIP export. |

## File-Level Expectations

| File | Expected outcome |
|---|---|
| `重复候选_人工智能产业发展报告.docx` | Should appear as a duplicate/alternative candidate or unused file. |
| `unrelated_无关材料.pdf` | Should remain unused. |
| `unreadable_corrupt.pdf` | Should produce a failure entry or extraction error without blocking other files. |

## Export Expectations

- `匹配清单.xlsx` is downloadable after matching completes, even when review is still needed.
- `发言材料.zip` is disabled until unresolved items are confirmed or skipped.
- After skipping orders 004 and 005 and handling any review item, ZIP should include ordered renamed files and `匹配清单.xlsx`.
- Expected ZIP names should start with order numbers such as `001_`, `002_`, and `003_` and preserve source extensions.
