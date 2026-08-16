## Context

See `proposal.md` for motivation and `specs/ordering-document-packaging/spec.md` for behavior requirements. The project already has document extraction, OCR, LangChain/Ollama calls, Streamlit UI, FastAPI routes, and in-memory task handling, but the ordering workflow needs a separate service layer because it combines Excel parsing, document identity extraction, matching, manual review, and ZIP export.

Current constraints that shape the design:

- New user-facing UI work must be integrated into `app_streamlit_v2.py`, not the legacy `app_streamlit.py`.
- Existing navigation already uses `⑤ Excel 合并`; the ordering workflow will be added as `⑥ 顺序编号` without replacing that page.
- Ollama must be used through `OLLAMA_BASE_URL` and must not be replaced by cloud model calls.
- Uploaded originals must not be renamed or overwritten.
- Initial storage can be in memory, but matching logic should remain independent from Streamlit so future API support can reuse it.
- Existing extraction code expects `extractor.extract(filename, raw, settings_dict)` while task settings are Pydantic objects.
- Existing temporary file cleanup currently conflicts with retry/export requirements.

## Goals / Non-Goals

**Goals:**

- Add a reusable ordering service module that Streamlit can call and future API routes can reuse.
- Keep deterministic matching rules independent from model calls.
- Use Ollama only for document identity extraction and limited ambiguous semantic decisions.
- Add Pydantic models that clearly represent order items, document candidates, matches, and ordering tasks.
- Add ZIP/report generation that is safe for Windows filenames, Chinese filenames, duplicate names, and path traversal.
- Add focused tests that mock model and OCR failure paths instead of calling paid or external services.

**Non-Goals:**

- No persistent database in the first version.
- No replacement of the existing template-based extraction workflow.
- No cloud LLM provider integration for ordering.
- No asynchronous job queue beyond current in-process execution.
- No automatic deletion of task files immediately after processing; cleanup remains explicit.
- No FastAPI ordering endpoints in this first implementation; API support can be added in a follow-up change.

## Decisions

### Decision: Put ordering business logic in `order_matching.py`

The ordering workflow will be implemented in a new module with pure functions and a small in-memory task store. Streamlit will call this module instead of embedding parsing, matching, or ZIP logic in the page.

Alternatives considered:

- Put logic directly in `app_streamlit_v2.py`: rejected because manual review would make the page hard to maintain and future API reuse would duplicate behavior.
- Extend `task_engine.py`: rejected because existing task engine is template-extraction focused and has different lifecycle semantics.

### Decision: Add ordering-specific Pydantic models to `schema.py`

`OrderItem`, `DocumentCandidate`, `MatchStatus`, `OrderMatch`, and `OrderingTask` will live with the existing models so serialization is consistent in Streamlit state and future API responses.

Alternatives considered:

- Dataclasses in `order_matching.py`: simpler internally but weaker for FastAPI response/request models.
- Separate schema module: cleaner separation but more churn for a small first version.

### Decision: Use deterministic matching before any semantic model call

The matcher will normalize codes and titles, compute deterministic scores, and assign unique matches. Ollama will only be used when candidates remain ambiguous after deterministic scoring and the candidate set is small.

Alternatives considered:

- Send all documents and order items to Ollama: rejected for latency, context size, reliability, and privacy risk.
- Use only filename matching: rejected because business rules require PDF/DOCX content extraction.

### Decision: Treat document identity extraction as layered

Each document candidate will be built from filename signals first, then extracted text from `extractor.extract()`, then Ollama JSON identity extraction when useful. If any layer fails, the candidate records the error and processing continues.

Alternatives considered:

- Always call Ollama first: rejected because filename/code matching is cheaper and more reliable.
- Require OCR success before matching: rejected because filename and DOCX/native text may still be enough.

### Decision: Store uploaded bytes for ordering tasks in memory for v1

Ordering tasks will keep original filenames, extensions, and bytes in memory for export and retry. This avoids modifying source uploads and avoids depending on temporary file availability for ZIP output.

Alternatives considered:

- Store only temporary file paths: fragile because cleanup can break retry/export.
- Add database/object storage: out of scope for first version.

### Decision: Fix existing task lifecycle separately from ordering task storage

The existing template-extraction `task_engine.py` will stop deleting temporary files at normal task completion and will expose explicit cleanup behavior. Ordering can use in-memory bytes while the existing workflow retains temporary files for retry/export.

Alternatives considered:

- Leave existing cleanup unchanged: rejected because it is a stated blocker.
- Disable cleanup globally without explicit cleanup: rejected because it risks permanent leaks.

### Decision: Generate ZIP and report from current match state only after validation

Export will validate duplicate assignments, unresolved matches, missing source bytes, and output filename conflicts before building the ZIP. File names will be sanitized and normalized, and duplicate output names will receive deterministic suffixes.

Alternatives considered:

- Export partial results silently: rejected because users need to know unresolved or conflicting assignments.
- Fail on duplicate sanitized names without suffixing: rejected because valid business cases can share titles.

### Decision: Defer FastAPI ordering endpoints

The first implementation will expose the ordering workflow through Streamlit and keep the core service functions reusable. FastAPI endpoints are deferred because the confirmed need is to add one feature to the existing front end, and a full API surface would increase scope without improving the immediate workflow.

Alternatives considered:

- Build Streamlit and FastAPI in one pass: rejected for this iteration because it expands testing and route design beyond the confirmed scope.
- Put ordering logic only in Streamlit: rejected because the service should remain reusable for future API support.

## Risks / Trade-offs

- [Risk] In-memory task storage loses data on process restart → Mitigation: document as v1 limitation and keep logic storage-independent for future persistence.
- [Risk] Large batches can consume memory when storing original bytes → Mitigation: keep initial concurrency low, expose file counts/sizes, and allow later swap to disk-backed storage.
- [Risk] Ollama 36B latency can be high → Mitigation: model concurrency defaults to 1, deterministic matching avoids unnecessary model calls, and timeouts default to 300 seconds.
- [Risk] OCR availability varies by machine → Mitigation: OCR failures are recorded per file and do not become valid model input.
- [Risk] Title matching can produce false positives → Mitigation: code has priority, auto-match requires score and margin thresholds, ambiguous results require manual review.
- [Risk] Temporary file cleanup changes may retain files longer than before → Mitigation: add explicit cleanup and tests for owned-file cleanup.
- [Risk] Proxy configuration can block internal Ollama access → Mitigation: do not blindly set `NO_PROXY`; report clear connection and timeout errors from the Ollama call path.

## Migration Plan

1. Add blocker fixes and tests first so existing extraction remains stable.
2. Add ordering models and `order_matching.py` with unit tests.
3. Add Streamlit page `⑥ 顺序编号` and verify app startup on configured port.
4. Run compile, import, `pip check`, pytest, and conflict-marker checks.

Rollback strategy: the new ordering workflow is additive. If needed, remove or hide the new Streamlit navigation entry while leaving existing extraction workflows intact.
