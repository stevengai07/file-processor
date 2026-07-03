# -*- coding: utf-8 -*-
"""
task_engine.py — In-memory task store and per-file extraction pipeline.

Responsibilities:
  - Create and store Task objects
  - Accept file uploads and queue them against a task
  - Run extraction pipeline concurrently (text -> AI -> validate -> status)
  - Expose per-file retry
  - Provide task / file status queries
  - Handle edits (PATCH) and auto-recalculate file status

MVP: all state is in-memory. Swap _TASK_STORE for a DB-backed store in production.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from schema import (
    ExtractionSettings,
    FieldIssue,
    FieldType,
    FieldValue,
    FileResult,
    FileStatus,
    IssueType,
    Task,
    TaskFile,
    TaskStatus,
    TemplateSnapshot,
)
from template_service import get_template


# ══════════════════════════════════════════════════════════════════════════════
# IN-MEMORY STORE
# ══════════════════════════════════════════════════════════════════════════════

_TASK_STORE: Dict[str, Task] = {}

SUPPORTED_EXTENSIONS = {".pdf", ".docx"}


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API — TASK LIFECYCLE
# ══════════════════════════════════════════════════════════════════════════════

def create_task(
    template_id: str,
    name: str = "未命名任务",
    settings: Optional[ExtractionSettings] = None,
) -> Task:
    """Create a new draft task linked to a confirmed template."""
    get_template(template_id)   # raises KeyError if not found
    task = Task(
        task_id=str(uuid.uuid4()),
        name=name,
        template_id=template_id,
        settings=settings or ExtractionSettings(),
        status=TaskStatus.DRAFT,
        created_at=datetime.utcnow(),
    )
    _TASK_STORE[task.task_id] = task
    return task


def add_files(task_id: str, file_tuples: List[Tuple[str, bytes]]) -> Dict[str, Any]:
    """
    Accept a list of (filename, file_bytes) tuples.
    Validates each file; saves accepted files to temp disk paths.
    Returns {accepted: [...], rejected: [...]}.
    """
    task = _get_task(task_id)
    if task.status != TaskStatus.DRAFT:
        raise ValueError("任务已启动，不允许继续添加文件。")

    accepted, rejected = [], []
    existing_sigs = {(f.filename, f.size) for f in task.files}

    for filename, file_bytes in file_tuples:
        ext = _file_ext(filename)
        if ext not in SUPPORTED_EXTENSIONS:
            rejected.append({
                "filename": filename,
                "reason": f"不支持的文件格式 {ext}，仅支持 .pdf 和 .docx。",
            })
            continue

        size = len(file_bytes)
        if (filename, size) in existing_sigs:
            rejected.append({"filename": filename, "reason": "同名同大小文件已存在，跳过。"})
            continue

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(file_bytes)
        tmp.flush()
        tmp.close()

        tf = TaskFile(
            file_id=str(uuid.uuid4()),
            filename=filename,
            size=size,
            file_type=ext.lstrip("."),
            status=FileStatus.PENDING,
            tmp_path=tmp.name,
        )
        task.files.append(tf)
        existing_sigs.add((filename, size))
        accepted.append({"file_id": tf.file_id, "filename": filename})

    return {"accepted": accepted, "rejected": rejected}


def start_task(task_id: str) -> Task:
    """Validate preconditions and launch background extraction."""
    task = _get_task(task_id)
    if task.status not in (TaskStatus.DRAFT, TaskStatus.QUEUED):
        raise ValueError(f"任务当前状态为 {task.status.value}，无法重复启动。")
    pending = [f for f in task.files if f.status == FileStatus.PENDING]
    if not pending:
        raise ValueError("没有可处理的文件。")
    template = get_template(task.template_id)
    task.status = TaskStatus.RUNNING
    task.started_at = datetime.utcnow()
    _run_extraction(task, template, pending)
    return task


def retry_file(task_id: str, file_id: str) -> TaskFile:
    """Reset a failed or cancelled file and re-run extraction."""
    task = _get_task(task_id)
    tf = task.get_file(file_id)
    if tf is None:
        raise KeyError(f"file_id {file_id!r} not found in task {task_id!r}")
    if tf.status not in (FileStatus.FAILED, FileStatus.CANCELLED):
        raise ValueError(f"只能重试状态为《失败》或《已取消》的文件，当前状态：{tf.status.value}")
    template = get_template(task.template_id)
    tf.status = FileStatus.PENDING
    existing = task.get_result(file_id)
    if existing:
        task.results.remove(existing)
    _process_one(task, template, tf)
    return tf


def cancel_task(task_id: str) -> Task:
    """Cancel all pending files in a task."""
    task = _get_task(task_id)
    for tf in task.files:
        if tf.status == FileStatus.PENDING:
            tf.status = FileStatus.CANCELLED
    if task.status in (TaskStatus.DRAFT, TaskStatus.QUEUED, TaskStatus.RUNNING):
        task.status = TaskStatus.CANCELLED
    return task


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API — QUERIES
# ══════════════════════════════════════════════════════════════════════════════

def get_task(task_id: str) -> Task:
    return _get_task(task_id)


def list_tasks() -> List[Task]:
    return sorted(_TASK_STORE.values(), key=lambda t: t.created_at, reverse=True)


def get_results(
    task_id: str,
    status: Optional[str] = None,
    keyword: Optional[str] = None,
    has_issues: bool = False,
    page: int = 1,
    page_size: int = 50,
) -> Dict[str, Any]:
    """Paginated, filtered result list."""
    task = _get_task(task_id)
    results = list(task.results)

    if status:
        results = [r for r in results if r.status.value == status]
    if keyword:
        kw = keyword.lower()
        results = [r for r in results if kw in r.filename.lower()]
    if has_issues:
        results = [r for r in results if r.issue_count > 0]

    total = len(results)
    start = (page - 1) * page_size
    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "results": results[start: start + page_size],
    }


def get_file_result(task_id: str, file_id: str) -> FileResult:
    task = _get_task(task_id)
    result = task.get_result(file_id)
    if result is None:
        raise KeyError(f"No result for file_id {file_id!r} in task {task_id!r}")
    return result


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API — EDIT
# ══════════════════════════════════════════════════════════════════════════════

def patch_result(
    task_id: str,
    file_id: str,
    field_updates: Dict[str, Any],
) -> FileResult:
    """
    Apply human edits to extracted field values.
    Validates each value against its template field type.
    Recalculates file status after applying all changes.
    """
    task = _get_task(task_id)
    result = task.get_result(file_id)
    if result is None:
        raise KeyError(f"No result for file_id {file_id!r}")
    template = get_template(task.template_id)

    for key, new_value in field_updates.items():
        fv = result.get_field(key)
        tfield = next((f for f in template.fields if f.key == key), None)
        if tfield is None:
            raise ValueError(f"字段 [{key}] 不在模板定义中。")

        type_err = _validate_type(new_value, tfield.type)
        if type_err:
            raise ValueError(f"字段 [{tfield.name}] 类型错误：{type_err}")

        if fv:
            fv.value = new_value
            fv.manually_edited = True
            fv.edited_at = datetime.utcnow()
        else:
            result.fields.append(FieldValue(
                key=key,
                name=tfield.name,
                value=new_value,
                raw_ai_value=None,
                manually_edited=True,
                edited_at=datetime.utcnow(),
            ))

        # Remove stale type-invalid issues for this field
        result.issues = [
            i for i in result.issues
            if not (i.field_key == key and i.issue_type == IssueType.TYPE_INVALID)
        ]

    result.recalculate_status(template)
    _update_task_status(task)
    return result


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACTION PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def _run_extraction(task: Task, template: TemplateSnapshot, files: List[TaskFile]) -> None:
    """Run extraction concurrently across all pending files."""
    with ThreadPoolExecutor(max_workers=task.settings.concurrency) as pool:
        futures = {pool.submit(_process_one, task, template, tf): tf for tf in files}
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as exc:
                tf = futures[future]
                tf.status = FileStatus.FAILED
                _upsert_result(task, FileResult(
                    file_id=tf.file_id,
                    filename=tf.filename,
                    status=FileStatus.FAILED,
                    error_message=str(exc),
                ))
    _update_task_status(task)
    _cleanup_tmp_files(task)


def _process_one(task: Task, template: TemplateSnapshot, tf: TaskFile) -> None:
    """
    Full pipeline for a single file:
      1. _extract_text  →  ExtractedDocument  (extractor.py)
      2. _run_ai        →  ExtractionResult   (agent.py)
      3. Merge fields + issues into FileResult
      4. Add any required-missing issues not already caught by agent
      5. Set file / task status
    """
    tf.status = FileStatus.PROCESSING
    prev = task.get_result(tf.file_id)
    retry_count = (prev.retry_count + 1) if prev else 0

    result = FileResult(
        file_id=tf.file_id,
        filename=tf.filename,
        status=FileStatus.PROCESSING,
        started_at=datetime.utcnow(),
        model_used=task.settings.model,
        retry_count=retry_count,
    )
    _upsert_result(task, result)

    try:
        # ── Step 1: text + OCR extraction ─────────────────────────────────
        doc = _extract_text(tf, task.settings)

        # ── Step 2: AI field extraction ───────────────────────────────────
        agent_result = _run_ai(doc, template, task.settings)

        # ── Step 3: merge into FileResult ─────────────────────────────────
        result.fields     = agent_result.fields
        result.issues     = list(agent_result.issues)
        result.model_used = agent_result.model_used
        result.retry_count = retry_count + agent_result.retry_count

        # ── Step 4: required-missing guard (belt-and-braces) ──────────────
        # agent.py already emits REQUIRED_MISSING issues; this catches any
        # that slipped through due to field-key mismatches or empty strings.
        existing_missing_keys = {
            i.field_key for i in result.issues
            if i.issue_type == IssueType.REQUIRED_MISSING
        }
        for k in template.required_keys:
            if k in existing_missing_keys:
                continue
            fv = result.get_field(k)
            if not fv or fv.value in (None, "", []):
                fdef = next((f for f in template.fields if f.key == k), None)
                label = fdef.name if fdef else k
                result.issues.append(FieldIssue(
                    field_key=k,
                    field_name=label,
                    issue_type=IssueType.REQUIRED_MISSING,
                    message=f"必填字段 [{label}] 未找到值",
                ))

        # ── Step 5: set status ────────────────────────────────────────────
        if result.issues:
            result.status = FileStatus.NEEDS_REVIEW
            tf.status     = FileStatus.NEEDS_REVIEW
        else:
            result.status = FileStatus.SUCCESS
            tf.status     = FileStatus.SUCCESS

    except Exception as e:
        result.status      = FileStatus.FAILED
        result.error_message = str(e)
        tf.status          = FileStatus.FAILED

    finally:
        result.completed_at = datetime.utcnow()
        _upsert_result(task, result)


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE STEPS — delegate to extractor.py and agent.py
# ══════════════════════════════════════════════════════════════════════════════

def _extract_text(tf: TaskFile, settings: ExtractionSettings):
    """
    Delegate to extractor.py.
    Returns an ExtractedDocument (not a plain string).

    extractor.extract(filename, raw_bytes, settings) handles:
      - PDF native text (pdfplumber)
      - PDF OCR fallback (Tesseract via image_enhancer.py)
      - DOCX paragraphs, tables, text boxes, headers/footers
    """
    from extractor import extract

    if not tf.tmp_path or not os.path.exists(tf.tmp_path):
        raise FileNotFoundError(f"临时文件不存在：{tf.tmp_path}")

    with open(tf.tmp_path, "rb") as fh:
        raw = fh.read()

    doc = extract(tf.filename, raw, settings)

    if not doc.total_chars:
        raise ValueError(
            "无法从文件中提取文本，请检查文件是否损坏或加密。"
        )

    return doc


def _run_ai(doc, template: TemplateSnapshot, settings: ExtractionSettings):
    """
    Delegate to agent.py.
    Returns an ExtractionResult with typed FieldValue objects and issues.

    agent.run_extraction() handles:
      - Prompt construction from TemplateField definitions
      - OpenAI / Anthropic routing via LangChain
      - JSON parsing + retry on parse failures
      - Type coercion (_to_int, _to_float, _to_date, _to_bool, _to_list)
      - source_snippet verification (anti-hallucination)
      - REQUIRED_MISSING and TYPE_INVALID issues
    """
    from agent import run_extraction

    return run_extraction(
        doc=doc,
        fields=template.fields,
        settings=settings,
        file_id=doc.filename,
    )


# ══════════════════════════════════════════════════════════════════════════════
# TYPE COERCION & VALIDATION  (used by patch_result only)
# ══════════════════════════════════════════════════════════════════════════════

def _coerce(value: Any, ftype: FieldType) -> Tuple[Any, Optional[str]]:
    """
    Try to coerce value to the expected FieldType.
    Returns (coerced_value, error_message_or_None).
    Used exclusively by patch_result to validate human edits.
    """
    if value is None or value == "":
        return None, None

    try:
        if ftype in (FieldType.TEXT, FieldType.LONG_TEXT):
            return str(value), None

        if ftype == FieldType.INTEGER:
            return int(float(str(value).replace(",", "").replace("，", ""))), None

        if ftype == FieldType.DECIMAL:
            return float(str(value).replace(",", "").replace("，", "")), None

        if ftype == FieldType.DATE:
            return _normalise_date(str(value)), None

        if ftype == FieldType.BOOLEAN:
            return _normalise_bool(value), None

        if ftype == FieldType.LIST:
            if isinstance(value, list):
                return [str(v) for v in value], None
            return [
                s.strip()
                for s in re.split(r"[,;\n，；]", str(value))
                if s.strip()
            ], None

    except Exception as e:
        return value, f"期望类型 {ftype.value}，实际内容无法转换：{e}"

    return value, None


def _validate_type(value: Any, ftype: FieldType) -> Optional[str]:
    """Return an error string if value cannot be coerced to ftype, else None."""
    _, err = _coerce(value, ftype)
    return err


def _normalise_date(s: str) -> str:
    """Best-effort date normalisation to YYYY-MM-DD."""
    s = s.strip()
    if re.match(r"\d{4}-\d{2}-\d{2}", s):
        return s
    # DD/MM/YYYY or MM/DD/YYYY
    m = re.match(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", s)
    if m:
        return f"{m.group(3)}-{m.group(1).zfill(2)}-{m.group(2).zfill(2)}"
    # YYYY年MM月DD日  or  YYYY/MM/DD  or  YYYY.MM.DD
    m = re.match(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
    return s


def _normalise_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    s = str(value).strip().lower()
    if s in ("true", "yes", "1", "是", "y", "是的", "t", "√", "✓", "有", "同意"):
        return True
    if s in ("false", "no", "0", "否", "n", "不是", "f", "×", "✗", "无", "不同意"):
        return False
    raise ValueError(f"无法识别布尔值：{value!r}")


# ══════════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _get_task(task_id: str) -> Task:
    if task_id not in _TASK_STORE:
        raise KeyError(f"Task {task_id!r} not found.")
    return _TASK_STORE[task_id]


def _upsert_result(task: Task, result: FileResult) -> None:
    """Insert or replace the FileResult for a given file_id."""
    task.results = [r for r in task.results if r.file_id != result.file_id]
    task.results.append(result)


def _update_task_status(task: Task) -> None:
    """Recalculate top-level task status from individual file statuses."""
    statuses = {f.status for f in task.files}
    if not statuses:
        return
    terminal = {
        FileStatus.SUCCESS,
        FileStatus.NEEDS_REVIEW,
        FileStatus.FAILED,
        FileStatus.CANCELLED,
    }
    if statuses <= terminal:
        task.status = TaskStatus.COMPLETED
        task.completed_at = datetime.utcnow()
    elif FileStatus.PROCESSING in statuses or FileStatus.PENDING in statuses:
        task.status = TaskStatus.RUNNING


def _cleanup_tmp_files(task: Task) -> None:
    """Delete temporary uploaded files after processing completes."""
    for tf in task.files:
        if tf.tmp_path and os.path.exists(tf.tmp_path):
            try:
                os.unlink(tf.tmp_path)
            except OSError:
                pass
        tf.tmp_path = None


def _file_ext(filename: str) -> str:
    return os.path.splitext(filename.lower())[1]