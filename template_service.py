# -*- coding: utf-8 -*-
"""
template_service.py — Excel template parsing, column auto-mapping, and snapshot storage.

Responsibilities:
  - Parse an uploaded .xlsx into a list of TemplateField objects
  - Auto-map column headers using known synonyms (Chinese + English)
  - Validate parsed fields before saving
  - Save confirmed templates as TemplateSnapshot objects (in-memory for MVP)
  - Retrieve snapshots by template_id
"""

from __future__ import annotations

import re
import uuid
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd

from schema import (
    FieldType,
    ParseTemplateResponse,
    TemplateField,
    TemplateSnapshot,
)


# ══════════════════════════════════════════════════════════════════════════════
# SYNONYM MAP
# Maps every recognised column alias -> canonical internal column role.
# ══════════════════════════════════════════════════════════════════════════════

_SYNONYM_MAP: Dict[str, str] = {
    # field name
    "字段名称": "name", "字段名": "name", "名称": "name",
    "field name": "name", "field": "name", "name": "name",
    "column name": "name", "列名": "name",
    # description
    "字段说明": "description", "说明": "description",
    "描述": "description", "备注": "description",
    "description": "description", "desc": "description",
    "field description": "description", "备注说明": "description",
    # data type
    "数据类型": "type", "类型": "type", "字段类型": "type",
    "data type": "type", "type": "type", "dtype": "type",
    "field type": "type",
    # required
    "是否必填": "required", "必填": "required", "required": "required",
    "is required": "required", "mandatory": "required", "必须": "required",
    # example
    "示例值": "example", "示例": "example", "example": "example",
    "sample": "example", "样例": "example", "例子": "example",
    "example value": "example",
    # prompt hint
    "提取提示": "prompt_hint", "提示": "prompt_hint",
    "提取说明": "prompt_hint", "同义词": "prompt_hint",
    "extraction hint": "prompt_hint", "hint": "prompt_hint",
    "prompt": "prompt_hint", "备注提示": "prompt_hint",
}

_TYPE_MAP: Dict[str, FieldType] = {
    "text": FieldType.TEXT, "文本": FieldType.TEXT, "字符串": FieldType.TEXT, "string": FieldType.TEXT, "str": FieldType.TEXT,
    "long_text": FieldType.LONG_TEXT, "长文本": FieldType.LONG_TEXT, "多行文本": FieldType.LONG_TEXT, "textarea": FieldType.LONG_TEXT,
    "integer": FieldType.INTEGER, "整数": FieldType.INTEGER, "int": FieldType.INTEGER, "number": FieldType.INTEGER,
    "decimal": FieldType.DECIMAL, "小数": FieldType.DECIMAL, "float": FieldType.DECIMAL, "numeric": FieldType.DECIMAL, "数字": FieldType.DECIMAL,
    "date": FieldType.DATE, "日期": FieldType.DATE, "datetime": FieldType.DATE, "时间": FieldType.DATE, "日期时间": FieldType.DATE,
    "boolean": FieldType.BOOLEAN, "布尔值": FieldType.BOOLEAN, "bool": FieldType.BOOLEAN, "是否": FieldType.BOOLEAN, "yes/no": FieldType.BOOLEAN,
    "list": FieldType.LIST, "列表": FieldType.LIST, "array": FieldType.LIST, "多值": FieldType.LIST, "multi": FieldType.LIST,
}

_TRUTHY = {"是", "yes", "true", "1", "y", "必填", "required", "√", "✓"}


# ══════════════════════════════════════════════════════════════════════════════
# IN-MEMORY STORES
# ══════════════════════════════════════════════════════════════════════════════

_UPLOAD_CACHE: Dict[str, dict] = {}
_TEMPLATE_STORE: Dict[str, TemplateSnapshot] = {}


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def parse_excel(
    file_bytes: bytes,
    filename: str,
    sheet_name: Optional[str] = None,
    header_row: int = 1,
) -> ParseTemplateResponse:
    """
    Parse an Excel file and return field previews.
    Caches raw bytes so switch_sheet works without re-upload.
    """
    import io
    wb = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = wb.sheet_names
    selected = sheet_name if sheet_name and sheet_name in sheets else sheets[0]
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=selected,
        header=header_row - 1,
        dtype=str,
    ).fillna("")
    upload_id = str(uuid.uuid4())
    _UPLOAD_CACHE[upload_id] = {
        "df": df,
        "sheets": sheets,
        "selected_sheet": selected,
        "filename": filename,
        "header_row": header_row,
        "raw_bytes": file_bytes,
    }
    fields, errors = _build_fields(df, filename)
    return ParseTemplateResponse(
        upload_id=upload_id,
        sheets=[str(s) for s in sheets],
        selected_sheet=str(selected),
        fields=fields,
        errors=errors,
    )


def switch_sheet(
    upload_id: str,
    sheet_name: str,
    header_row: int = 1,
) -> ParseTemplateResponse:
    """Re-parse the already-uploaded file using a different sheet."""
    import io
    if upload_id not in _UPLOAD_CACHE:
        raise KeyError(f"upload_id {upload_id!r} not found. Please re-upload the file.")
    cache = _UPLOAD_CACHE[upload_id]
    raw = cache.get("raw_bytes")
    if raw is None:
        raise ValueError("Raw bytes not cached. Call parse_excel again.")
    cache["selected_sheet"] = sheet_name
    cache["header_row"] = header_row
    df = pd.read_excel(
        io.BytesIO(raw),
        sheet_name=sheet_name,
        header=header_row - 1,
        dtype=str,
    ).fillna("")
    cache["df"] = df
    fields, errors = _build_fields(df, cache["filename"])
    return ParseTemplateResponse(
        upload_id=upload_id,
        sheets=[str(s) for s in cache["sheets"]],
        selected_sheet=str(sheet_name),
        fields=fields,
        errors=errors,
    )


def save_template(
    upload_id: str,
    fields: List[TemplateField],
    name: str = "未命名模板",
) -> TemplateSnapshot:
    """Validate and persist a confirmed template snapshot."""
    cache = _UPLOAD_CACHE.get(upload_id, {})
    source_file = cache.get("filename")
    snapshot = TemplateSnapshot(
        template_id=str(uuid.uuid4()),
        name=name,
        version=1,
        fields=fields,
        created_at=datetime.utcnow(),
        source_file=source_file,
    )
    _TEMPLATE_STORE[snapshot.template_id] = snapshot
    return snapshot


def get_template(template_id: str) -> TemplateSnapshot:
    """Retrieve a saved template snapshot by ID."""
    if template_id not in _TEMPLATE_STORE:
        raise KeyError(f"Template {template_id!r} not found.")
    return _TEMPLATE_STORE[template_id]


def list_templates() -> List[TemplateSnapshot]:
    """Return all saved templates, newest first."""
    return sorted(_TEMPLATE_STORE.values(), key=lambda t: t.created_at, reverse=True)


# ══════════════════════════════════════════════════════════════════════════════
# INTERNAL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _normalise(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _map_columns(columns: List[str]) -> Tuple[Dict[str, str], List[str]]:
    """Map raw column headers to canonical roles using _SYNONYM_MAP."""
    col_role: Dict[str, str] = {}
    unmapped: List[str] = []
    assigned_roles: set = set()
    for col in columns:
        norm = _normalise(col)
        role = _SYNONYM_MAP.get(norm)
        if role and role not in assigned_roles:
            col_role[col] = role
            assigned_roles.add(role)
        else:
            unmapped.append(col)
    return col_role, unmapped


def _parse_type(raw: str) -> FieldType:
    return _TYPE_MAP.get(_normalise(raw), FieldType.TEXT)


def _parse_required(raw: str) -> bool:
    return _normalise(raw) in _TRUTHY


def _make_key(name: str) -> str:
    key = re.sub(r"[^\w\u4e00-\u9fff]+", "_", name.strip().lower()).strip("_")
    return key or "field"


def _get(row: pd.Series, role_col: Dict[str, str], role: str) -> str:
    col = role_col.get(role)
    if col is None:
        return ""
    return str(row.get(col, "")).strip()


def _build_fields(df: pd.DataFrame, filename: str) -> Tuple[List[TemplateField], List[str]]:
    """Convert a parsed DataFrame into TemplateField objects."""
    errors: List[str] = []
    col_role, _ = _map_columns(list(df.columns))
    role_col = {v: k for k, v in col_role.items()}

    if "name" not in role_col:
        errors.append(
            "未能识别字段名称列。请确保模板包含[字段名称]（或同义列名）列。"
            f"当前列名：{list(df.columns)}"
        )
        return [], errors

    name_col = role_col["name"]
    fields: List[TemplateField] = []
    seen_names: set = set()

    for idx, row in df.iterrows():
        raw_name = str(row.get(name_col, "")).strip()
        if not raw_name:
            continue
        row_num = int(idx) + 2
        if raw_name in seen_names:
            errors.append(f"第 {row_num} 行：字段名称 [{raw_name}] 重复，已跳过。")
            continue
        seen_names.add(raw_name)
        key         = _make_key(raw_name)
        description = _get(row, role_col, "description")
        raw_type    = _get(row, role_col, "type")
        raw_req     = _get(row, role_col, "required")
        example     = _get(row, role_col, "example")
        prompt_hint = _get(row, role_col, "prompt_hint")
        field_type  = _parse_type(raw_type) if raw_type else FieldType.TEXT
        required    = _parse_required(raw_req) if raw_req else False
        try:
            fields.append(TemplateField(
                key=key,
                name=raw_name,
                description=description or None,
                type=field_type,
                required=required,
                example=example or None,
                prompt_hint=prompt_hint or None,
                order=len(fields),
            ))
        except Exception as e:
            errors.append(f"第 {row_num} 行（{raw_name}）解析失败：{e}")

    if not fields:
        errors.append("模板中未找到任何有效字段，请检查文件内容。")
    return fields, errors


# ══════════════════════════════════════════════════════════════════════════════
# CLI smoke-test: python template_service.py sample_template.xlsx
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "sample_template.xlsx"
    with open(path, "rb") as fh:
        data = fh.read()
    resp = parse_excel(data, path)
    print(f"Sheets : {resp.sheets}")
    print(f"Active : {resp.selected_sheet}")
    print(f"Fields : {len(resp.fields)}")
    for f in resp.fields:
        req = "[必填]" if f.required else "      "
        print(f"  {req} {f.order+1:2}. {f.name:<20} ({f.type.value})")
    if resp.errors:
        print("\nErrors:")
        for e in resp.errors:
            print(f"  ! {e}")
    snap = save_template(resp.upload_id, resp.fields, name="测试模板")
    print(f"\nSaved  : template_id={snap.template_id}")
    print(f"Keys   : {snap.field_keys}")
    print(f"Required: {snap.required_keys}")