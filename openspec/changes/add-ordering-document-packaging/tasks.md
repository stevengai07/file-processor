## 1. Existing Blocker Fixes

- [x] 1.1 Add `max_tokens` and `timeout_seconds` to `ExtractionSettings` and default the model to `qwen3.6:latest`.
- [x] 1.2 Update `agent.py` so Ollama calls use environment-driven `OLLAMA_BASE_URL`, consistent timeouts, and correct `model_used` values.
- [x] 1.3 Update `task_engine.py` to convert `ExtractionSettings` into an extractor settings dictionary before calling `extractor.extract()`.
- [x] 1.4 Change existing task temporary-file lifecycle so files remain available for retry/export until explicit cleanup.
- [x] 1.5 Ensure OCR disabled mode truly disables OCR fallback.
- [x] 1.6 Ensure OCR failure marker text is not treated as valid document body text for model extraction.
- [ ] 1.7 Add focused tests for settings compatibility, task extraction settings mapping, OCR disabled behavior, OCR failure handling, retry availability, and cleanup.

## 2. Ordering Models

- [x] 2.1 Add `OrderItem`, `DocumentCandidate`, `MatchStatus`, `OrderMatch`, and `OrderingTask` Pydantic models to `schema.py`.
- [x] 2.2 Add helper validation or methods needed for duplicate assignment and task KPI calculation.
- [x] 2.3 Add tests for ordering model serialization, default values, and status transitions.

## 3. Ordering Service Module

- [x] 3.1 Create `order_matching.py` with in-memory ordering task storage and storage-independent pure functions.
- [x] 3.2 Implement `parse_order_excel()` with sheet selection, header row selection, alias-based column mapping, manual mapping support, empty-row skipping, duplicate-order detection, invalid-row reporting, and source-row preservation.
- [x] 3.3 Implement `normalise_code()` and `normalise_title()` according to the normalization rules.
- [x] 3.4 Implement filename identity extraction for possible title and material code signals.
- [x] 3.5 Implement `extract_document_identity()` using `extractor.extract()` plus Ollama JSON extraction with per-file failure isolation and result caching.
- [ ] 3.6 Implement Ollama ambiguous-candidate resolution with bounded candidate input, environment-driven base URL, timeout handling, and invalid JSON handling.
- [x] 3.7 Implement `calculate_match_score()` with exact code, normalized code, exact title, containment, title similarity, Ollama-confirmed, and low-keyword-score outcomes.
- [x] 3.8 Implement `assign_unique_matches()` and `resolve_match()` with global one-to-one constraints, score threshold, margin threshold, occupied-file detection, and review status assignment.
- [x] 3.9 Implement manual confirmation, reassignment, skip handling, and conflict validation helpers.
- [x] 3.10 Implement safe output filename generation for illegal characters, leading/trailing spaces and dots, long names, duplicate names, Chinese names, extension preservation, and path traversal prevention.
- [x] 3.11 Implement `build_matching_report()` producing `匹配清单.xlsx` with all required columns.
- [x] 3.12 Implement `build_renamed_zip()` producing ordered renamed files plus the matching report without modifying source bytes or filenames.

## 4. Ordering Service Tests

- [x] 4.1 Add Excel parsing tests for alias mapping, manual mapping, duplicate order numbers, invalid rows, sheet selection, and header-row selection.
- [x] 4.2 Add normalization and scoring tests for exact code, normalized code, exact title, title containment, title similarity, no candidates, and multiple candidates.
- [x] 4.3 Add assignment tests for automatic unique assignment, low-confidence review, multi-candidate review, duplicate-file prevention, manual confirmation, and manual reassignment.
- [x] 4.4 Add filename tests for illegal characters, duplicate output names, long names, extension preservation, and Chinese filenames.
- [x] 4.5 Add ZIP/report tests for ZIP content completeness, `匹配清单.xlsx` presence, Chinese filenames, report columns, and source byte immutability.
- [ ] 4.6 Add mocked failure tests for Ollama unavailable, Ollama invalid JSON, OCR unavailable, single-file failure isolation, retry availability, and cleanup timing.

## 5. Streamlit UI

- [x] 5.1 Add navigation entry `⑥ 顺序编号` to `app_streamlit_v2.py` without changing the existing `⑤ Excel 合并` page or the legacy entry point.
- [x] 5.2 Build the order Excel upload UI with worksheet selection, header-row selection, automatic mapping preview, manual mapping fallback, row validation, and confirmation.
- [x] 5.3 Build PDF/DOCX batch upload UI showing file count and total size.
- [x] 5.4 Add OCR language and image enhancement controls using existing sidebar/session settings where appropriate.
- [x] 5.5 Add start extraction and matching action with progress display and cached document identity results.
- [x] 5.6 Add KPI cards for order count, input file count, matched, needs review, unmatched, unused files, and failures.
- [x] 5.7 Add automatic match result table.
- [x] 5.8 Add manual review UI showing order details, candidate file details, score, reason, source snippet, occupancy status, selection control, confirm action, and skip action.
- [x] 5.9 Add pre-export validation for duplicate assignments and unresolved items.
- [x] 5.10 Add ZIP download and report download controls.
- [ ] 5.11 Verify `app_streamlit_v2.py` starts successfully on the configured Streamlit port.

## 6. Final Verification

- [x] 6.1 Run the full project `py_compile` command including `order_matching.py`.
- [x] 6.2 Run the import smoke test for `agent`, `app_fastapi`, `extractor`, `schema`, and `order_matching`.
- [x] 6.3 Run `.venv/bin/python -m pip check`.
- [x] 6.4 Run pytest.
- [x] 6.5 Check for unresolved conflict markers.
- [ ] 6.6 Manually verify Streamlit ordering workflow with a small sample Excel and PDF/DOCX set.
- [ ] 6.7 Manually verify existing FastAPI `/health` and `/docs` still start, without adding ordering routes in this change.
