# -*- coding: utf-8 -*-
"""
schema.py — Pydantic models for the AI Document Extraction System.

Sections:
  1. Legacy fixed-schema model (ExtractedInfo) — kept for backward compatibility
  2. Enums
  3. Template models       — TemplateField, TemplateSnapshot
  4. Task & file models    — TaskFile, Task
  5. Extraction result     — FieldValue, FileResult
  6. Export models         — ExportRequest
  7. API request/response  — ParseTemplateResponse, ErrorDetail
"""

from __future__ import annotations
from datetime import datetime
from enum import Enum
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field, field_validator, model_validator


# ══════════════════════════════════════════════════════════════════════════════
# 1. LEGACY MODEL — backward compatible with existing agent.py / app_streamlit.py
# ══════════════════════════════════════════════════════════════════════════════

class ExtractedInfo(BaseModel):
    title:        str                 = Field(description="Title or subject of the document")
    summary:      str                 = Field(description="2-4 sentence summary of the document")
    key_topics:   List[str]           = Field(description="Main topics or themes discussed")
    entities:     List[str]           = Field(description="People, organizations, or places mentioned")
    dates:        List[str]           = Field(description="Important dates or time references found")
    action_items: Optional[List[str]] = Field(default=None, description="Action items, tasks, or recommendations")
    sentiment:    Optional[str]       = Field(default=None, description="Overall tone: positive, neutral, or negative")


# ══════════════════════════════════════════════════════════════════════════════
# 2. ENUMS
# ══════════════════════════════════════════════════════════════════════════════

class FieldType(str, Enum):
    TEXT      = "text"
    LONG_TEXT = "long_text"
    INTEGER   = "integer"
    DECIMAL   = "decimal"
    DATE      = "date"
    BOOLEAN   = "boolean"
    LIST      = "list"


class TaskStatus(str, Enum):
    DRAFT     = "draft"
    QUEUED    = "queued"
    RUNNING   = "running"
    COMPLETED = "completed"
    CANCELLED = "cancelled"
    FAILED    = "failed"


class FileStatus(str, Enum):
    PENDING      = "pending"
    PROCESSING   = "processing"
    SUCCESS      = "success"
    NEEDS_REVIEW = "needs_review"
    FAILED       = "failed"
    CANCELLED    = "cancelled"


class IssueType(str, Enum):
    REQUIRED_MISSING = "required_missing"
    TYPE_INVALID     = "type_invalid"
    VALUE_AMBIGUOUS  = "value_ambiguous"
    SOURCE_NOT_FOUND = "source_not_found"
    EXTRACTION_ERROR = "extraction_error"
    OCR_ERROR        = "ocr_error"


class ExportFormat(str, Enum):
    EXCEL = "excel"
    DOCX  = "docx"


class ExportScope(str, Enum):
    ALL          = "all"
    SUCCESS_ONLY = "success_only"
    SELECTED     = "selected"


# ══════════════════════════════════════════════════════════════════════════════
# 3. TEMPLATE MODELS
# ══════════════════════════════════════════════════════════════════════════════

class TemplateField(BaseModel):
    """One extractable field as defined by the user's Excel template."""
    key:         str            = Field(description="Stable snake_case identifier, auto-derived from name")
    name:        str            = Field(description="Display name shown in UI and used as column header")
    description: Optional[str] = Field(default=None, description="Business meaning, helps AI understand the field")
    type:        FieldType      = Field(default=FieldType.TEXT, description="Expected data type for validation")
    required:    bool           = Field(default=False, description="Missing value triggers needs_review status")
    example:     Optional[str]  = Field(default=None, description="Sample value to guide the AI")
    prompt_hint: Optional[str]  = Field(default=None, description="Synonyms, location hints, format hints for AI")
    order:       int            = Field(default=0, description="Column order in results table and export")

    @field_validator("name")
    @classmethod
    def name_not_empty(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("字段名称不能为空")
        return v

    @field_validator("key", mode="before")
    @classmethod
    def derive_key(cls, v: str) -> str:
        import re
        v = str(v).strip().lower()
        v = re.sub(r"[^\w\u4e00-\u9fff]+", "_", v)
        v = v.strip("_")
        return v or "field"


class TemplateSnapshot(BaseModel):
    """
    Immutable snapshot saved when user confirms a template.
    Tasks reference this snapshot so later edits do not affect running tasks.
    """
    template_id: str
    name:        str                 = Field(default="未命名模板")
    version:     int                 = Field(default=1)
    fields:      List[TemplateField]
    created_at:  datetime            = Field(default_factory=datetime.utcnow)
    source_file: Optional[str]       = Field(default=None, description="Original Excel filename")

    @model_validator(mode="after")
    def validate_fields(self) -> "TemplateSnapshot":
        if not self.fields:
            raise ValueError("模板至少包含一个有效字段")
        names = [f.name for f in self.fields]
        if len(names) != len(set(names)):
            raise ValueError("字段名称在同一模板内必须唯一")
        return self

    @property
    def field_keys(self) -> List[str]:
        return [f.key for f in sorted(self.fields, key=lambda x: x.order)]

    @property
    def required_keys(self) -> List[str]:
        return [f.key for f in self.fields if f.required]


# ══════════════════════════════════════════════════════════════════════════════
# 4. TASK & FILE MODELS
# ══════════════════════════════════════════════════════════════════════════════

@property
def counts(self) -> Dict[str, int]:
    from collections import Counter
    c = Counter(f.status.value for f in self.files)
    return {"total": len(self.files), "pending": c.get("pending",0),
            "processing": c.get("processing",0), "success": c.get("success",0),
            "needs_review": c.get("needs_review",0), "failed": c.get("failed",0),
            "cancelled": c.get("cancelled",0)}

@property
def progress_pct(self) -> float:
    if not self.files: return 0.0
    terminal = {"success","needs_review","failed","cancelled"}
    done = sum(1 for f in self.files if f.status.value in terminal)
    return round(done / len(self.files) * 100, 1)

def get_file(self, file_id: str):
    return next((f for f in self.files if f.file_id == file_id), None)

def get_result(self, file_id: str):
    return next((r for r in self.results if r.file_id == file_id), None)

class ExtractionSettings(BaseModel):
    """Processing parameters chosen by the user on Page 2."""
    model:       str  = Field(default="gpt-4o-mini")
    ocr_preset:  str  = Field(default="scanner")
    ocr_lang:    str  = Field(default="eng")
    mixed_eng:   bool = Field(default=False, description="Append +eng to OCR lang for mixed-language docs")
    concurrency: int  = Field(default=4, ge=1, le=20)

    @property
    def effective_lang(self) -> str:
        if self.mixed_eng and "eng" not in self.ocr_lang:
            return self.ocr_lang + "+eng"
        return self.ocr_lang


class FieldIssue(BaseModel):
    """One validation problem on a specific field within a file result."""
    field_key:  str
    field_name: str
    issue_type: IssueType
    message:    str
    raw_value:  Optional[Any] = None


class FieldValue(BaseModel):
    """The extracted (and optionally human-edited) value for one template field."""
    key:             str
    name:            str
    value:           Optional[Any]      = None
    raw_ai_value:    Optional[Any]      = None
    manually_edited: bool               = False
    edited_at:       Optional[datetime] = None
    source_snippet:  Optional[str]      = None


class FileResult(BaseModel):
    """All extraction results for one file within a task."""
    file_id:       str
    filename:      str
    status:        FileStatus         = FileStatus.PENDING
    fields:        List[FieldValue]   = Field(default_factory=list)
    issues:        List[FieldIssue]   = Field(default_factory=list)
    error_message: Optional[str]      = None
    started_at:    Optional[datetime] = None
    completed_at:  Optional[datetime] = None
    model_used:    Optional[str]      = None
    retry_count:   int                = 0

    @property
    def elapsed_seconds(self) -> Optional[float]:
        if self.started_at and self.completed_at:
            return (self.completed_at - self.started_at).total_seconds()
        return None

    @property
    def issue_count(self) -> int:
        return len(self.issues)

    def get_field(self, key: str) -> Optional[FieldValue]:
        return next((f for f in self.fields if f.key == key), None)

    def recalculate_status(self, template: TemplateSnapshot) -> None:
        """Re-evaluate file status after a human edit."""
        self.issues = [i for i in self.issues if i.issue_type != IssueType.REQUIRED_MISSING]
        missing_required = [
            k for k in template.required_keys
            if not self.get_field(k) or self.get_field(k).value in (None, "", [])
        ]
        for k in missing_required:
            field = next((f for f in template.fields if f.key == k), None)
            label = field.name if field else k
            self.issues.append(FieldIssue(
                field_key=k,
                field_name=label,
                issue_type=IssueType.REQUIRED_MISSING,
                message="必填字段 [" + label + "] 未找到值",
            ))
        if missing_required or any(i.issue_type == IssueType.TYPE_INVALID for i in self.issues):
            self.status = FileStatus.NEEDS_REVIEW
        else:
            self.status = FileStatus.SUCCESS


class TaskFile(BaseModel):
    """Lightweight file record stored in the task before processing begins."""
    file_id:   str
    filename:  str
    size:      int           = 0
    file_type: str           = "pdf"
    status:    FileStatus    = FileStatus.PENDING
    tmp_path:  Optional[str] = None


class Task(BaseModel):
    """Top-level task object — holds all files and links to a template snapshot."""
    task_id:      str
    name:         str                 = Field(default="未命名任务")
    template_id:  str
    settings:     ExtractionSettings  = Field(default_factory=ExtractionSettings)
    status:       TaskStatus          = TaskStatus.DRAFT
    files:        List[TaskFile]      = Field(default_factory=list)
    results:      List[FileResult]    = Field(default_factory=list)
    created_at:   datetime            = Field(default_factory=datetime.utcnow)
    started_at:   Optional[datetime]  = None
    completed_at: Optional[datetime]  = None

    @property
    def total(self) -> int:
        return len(self.files)

    @property
    def counts(self) -> Dict[str, int]:
        statuses = [f.status for f in self.files]
        return {
            "total":        len(statuses),
            "pending":      statuses.count(FileStatus.PENDING),
            "processing":   statuses.count(FileStatus.PROCESSING),
            "success":      statuses.count(FileStatus.SUCCESS),
            "needs_review": statuses.count(FileStatus.NEEDS_REVIEW),
            "failed":       statuses.count(FileStatus.FAILED),
            "cancelled":    statuses.count(FileStatus.CANCELLED),
        }

    @property
    def progress_pct(self) -> float:
        if not self.files:
            return 0.0
        terminal = (FileStatus.SUCCESS, FileStatus.NEEDS_REVIEW,
                    FileStatus.FAILED, FileStatus.CANCELLED)
        done = sum(1 for f in self.files if f.status in terminal)
        return round(done / len(self.files) * 100, 1)

    def get_result(self, file_id: str) -> Optional[FileResult]:
        return next((r for r in self.results if r.file_id == file_id), None)

    def get_file(self, file_id: str) -> Optional[TaskFile]:
        return next((f for f in self.files if f.file_id == file_id), None)


# ══════════════════════════════════════════════════════════════════════════════
# 5. EXPORT MODELS
# ══════════════════════════════════════════════════════════════════════════════

class ExportRequest(BaseModel):
    format:      ExportFormat        = ExportFormat.EXCEL
    scope:       ExportScope         = ExportScope.SUCCESS_ONLY
    file_ids:    Optional[List[str]] = None
    include_log: bool                = True


# ══════════════════════════════════════════════════════════════════════════════
# 6. API REQUEST / RESPONSE MODELS
# ══════════════════════════════════════════════════════════════════════════════

class ParseTemplateRequest(BaseModel):
    sheet_name: Optional[str] = None
    header_row: int           = Field(default=1, ge=1)


class ParseTemplateResponse(BaseModel):
    upload_id:      str
    sheets:         List[str]
    selected_sheet: str
    fields:         List[TemplateField]
    errors:         List[str] = Field(default_factory=list)


class SaveTemplateRequest(BaseModel):
    upload_id:  str
    name:       str                = "未命名模板"
    fields:     List[TemplateField]
    sheet_name: Optional[str]      = None
    header_row: int                = 1


class CreateTaskRequest(BaseModel):
    template_id: str
    name:        str                = "未命名任务"
    settings:    ExtractionSettings = Field(default_factory=ExtractionSettings)


class PatchResultRequest(BaseModel):
    """PATCH body — only the changed field key/value pairs."""
    fields: Dict[str, Any] = Field(description="Map of field_key to new value")


class ErrorDetail(BaseModel):
    code:    str
    message: str
    details: Optional[Dict[str, Any]] = None


class ParseTemplateResponse(BaseModel):
    upload_id: str; sheets: List[str]; selected_sheet: str
    fields: List[TemplateField]; errors: List[str] = []

class SaveTemplateRequest(BaseModel):
    upload_id: str; name: str = "未命名模板"; fields: List[TemplateField]

class CreateTaskRequest(BaseModel):
    template_id: str; name: str = "未命名任务"
    settings: Optional[ExtractionSettings] = None

class PatchResultRequest(BaseModel):
    fields: Dict[str, Any]

class ExportRequest(BaseModel):
    scope: ExportScope = ExportScope.SUCCESS_ONLY
    file_ids: Optional[List[str]] = None
    include_log: bool = True

class ErrorDetail(BaseModel):
    code: str; message: str; details: Optional[Dict[str, Any]] = None