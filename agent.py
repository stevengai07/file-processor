# -*- coding: utf-8 -*-
"""
agent.py — LangChain-powered field extraction agent.

Responsibilities:
  1. Receive an ExtractedDocument + list of TemplateFields.
  2. Build a structured prompt asking the LLM to locate and extract each field.
  3. Parse and type-cast the LLM response into typed FieldValue objects.
  4. Validate every field against its type + required constraint.
  5. Return an ExtractionResult with values, issues, confidence, and source snippets.

Strategy:
  - Single structured JSON call (function-calling / tool-use) per document.
  - For very long documents, split into chunks and merge results.
  - Retry with narrowed context on parse failures (max 2 retries).
  - Supports OpenAI and Anthropic models via LangChain.

Public API:
  run_extraction(doc, fields, settings) -> ExtractionResult
"""

from __future__ import annotations

import json
import logging
import re
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple

from schema import (
    ExtractionSettings,
    FieldType,
    FieldValue,
    IssueType,
    ResultIssue,
    TemplateField,
)
from extractor import ExtractedDocument

log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# RESULT TYPES
# ══════════════════════════════════════════════════════════════════════════════

class ExtractionResult:
    """
    Holds all extracted field values and validation issues for one document.
    """

    def __init__(
        self,
        file_id:      str,
        filename:     str,
        fields:       List[FieldValue],
        issues:       List[ResultIssue],
        model_used:   str,
        elapsed_seconds: float,
        retry_count:  int = 0,
    ):
        self.file_id          = file_id
        self.filename         = filename
        self.fields           = fields
        self.issues           = issues
        self.model_used       = model_used
        self.elapsed_seconds  = elapsed_seconds
        self.retry_count      = retry_count

    @property
    def issue_count(self) -> int:
        return len(self.issues)

    @property
    def has_issues(self) -> bool:
        return len(self.issues) > 0

    def get_field(self, key: str) -> Optional[FieldValue]:
        return next((f for f in self.fields if f.key == key), None)


# ══════════════════════════════════════════════════════════════════════════════
# PUBLIC API
# ══════════════════════════════════════════════════════════════════════════════

def run_extraction(
    doc:      ExtractedDocument,
    fields:   List[TemplateField],
    settings: Optional[ExtractionSettings] = None,
    file_id:  str = "",
) -> ExtractionResult:
    """
    Main entry point.  Build prompt → call LLM → parse → validate → return.

    Raises:
        AgentError: If the LLM cannot be reached or all retries fail.
    """
    import time

    if settings is None:
        settings = ExtractionSettings()

    t_start = time.perf_counter()

    ordered_fields = sorted(fields, key=lambda f: f.order)
    text = _prepare_text(doc, settings)

    raw_json: Optional[dict] = None
    retry_count = 0
    last_error: Optional[Exception] = None

    for attempt in range(3):
        try:
            raw_json = _call_llm(text, ordered_fields, settings)
            break
        except _ParseError as e:
            log.warning("Attempt %d parse error: %s", attempt + 1, e)
            last_error = e
            retry_count += 1
            # Narrow context: retry with just the first 6000 chars
            if attempt == 1:
                text = text[:6000]
        except Exception as e:
            log.error("Attempt %d LLM error: %s", attempt + 1, e)
            last_error = e
            retry_count += 1

    elapsed = time.perf_counter() - t_start

    if raw_json is None:
        raise AgentError(
            f"提取失败，已重试 {retry_count} 次。最后错误：{last_error}"
        )

    field_values, issues = _parse_and_validate(raw_json, ordered_fields, doc)

    return ExtractionResult(
        file_id=file_id or doc.filename,
        filename=doc.filename,
        fields=field_values,
        issues=issues,
        model_used=settings.model,
        elapsed_seconds=round(elapsed, 2),
        retry_count=retry_count,
    )


# ══════════════════════════════════════════════════════════════════════════════
# TEXT PREPARATION
# ══════════════════════════════════════════════════════════════════════════════

# Maximum characters to feed into a single prompt.
# GPT-4o context window ≈ 128k tokens; at ~3 chars/token that's ~384k chars.
# We stay well under to leave room for the prompt scaffold and JSON response.
_MAX_CHARS_SINGLE = 80_000
_MAX_CHARS_CHUNK  = 12_000   # size of each chunk for long documents
_CHUNK_OVERLAP    = 500      # overlap to avoid splitting mid-sentence


def _prepare_text(doc: ExtractedDocument, settings: ExtractionSettings) -> str:
    """
    Produce the document text for the prompt.
    Short docs: full text.  Long docs: chunked with overlap markers.
    """
    full = doc.full_text
    if len(full) <= _MAX_CHARS_SINGLE:
        return full

    # Document is very long — build a condensed representation
    log.debug("Document %s is long (%d chars) — chunking", doc.filename, len(full))
    chunks = _chunk_text(full, _MAX_CHARS_CHUNK, _CHUNK_OVERLAP)
    # Truncate to first N chunks that fit in budget
    budget = _MAX_CHARS_SINGLE
    parts: List[str] = []
    used = 0
    for i, chunk in enumerate(chunks):
        if used + len(chunk) > budget:
            parts.append(f"\n[… 文档过长，已截断，共 {doc.total_chars:,} 字符 …]")
            break
        parts.append(f"[片段 {i+1}]\n{chunk}")
        used += len(chunk)
    return "\n\n".join(parts)


def _chunk_text(text: str, size: int, overlap: int) -> List[str]:
    chunks: List[str] = []
    start = 0
    while start < len(text):
        end = min(start + size, len(text))
        chunks.append(text[start:end])
        start += size - overlap
    return chunks


# ══════════════════════════════════════════════════════════════════════════════
# PROMPT CONSTRUCTION
# ══════════════════════════════════════════════════════════════════════════════

_SYSTEM_PROMPT = """\
你是一位专业的文档信息提取助手。
你的任务是从用户提供的文档文本中，精确提取指定字段的值。

规则：
1. 严格按照字段定义提取，不要推测或填入文档中没有的内容。
2. 如果某字段确实找不到对应内容，返回 null，不要编造。
3. 日期统一格式化为 YYYY-MM-DD。
4. 数字只返回数值，不含货币符号或单位（除非字段定义要求）。
5. 列表类字段返回 JSON 数组。
6. 布尔类字段：是/有/同意/true → true，否/无/不同意/false → false。
7. 长文本字段摘录原文，不要二次改写。
8. 对每个字段同时返回 source_snippet：文档中支持该答案的原始片段（≤150字）。
9. 以 JSON 格式返回结果，结构如下所示。
"""

_RESPONSE_SCHEMA_TEMPLATE = """\
返回格式（严格 JSON，不含 Markdown 代码块）：
{{
{field_lines}
}}

每个字段的结构：
{{
  "value": <提取的值 或 null>,
  "source_snippet": "<文档中的原文片段>"
}}
"""


def _build_field_lines(fields: List[TemplateField]) -> str:
    lines = []
    for f in fields:
        type_hint = _type_hint(f.type)
        required_note = "【必填】" if f.required else "【可选】"
        example_note  = f"  示例：{f.example}" if f.example else ""
        hint_note     = f"  提示：{f.prompt_hint}" if f.prompt_hint else ""
        lines.append(
            f'  "{f.key}": {{'
            f'"value": <{type_hint}>{required_note}{example_note}{hint_note}, '
            f'"source_snippet": "<原文片段>"}}'
        )
    return ",\n".join(lines)


def _type_hint(ftype: FieldType) -> str:
    return {
        FieldType.TEXT:      "string",
        FieldType.LONG_TEXT: "string（长文本）",
        FieldType.INTEGER:   "integer",
        FieldType.DECIMAL:   "number",
        FieldType.DATE:      "YYYY-MM-DD string",
        FieldType.BOOLEAN:   "true/false",
        FieldType.LIST:      "array of strings",
    }.get(ftype, "string")


def _build_user_prompt(text: str, fields: List[TemplateField]) -> str:
    field_lines = _build_field_lines(fields)
    schema_block = _RESPONSE_SCHEMA_TEMPLATE.format(field_lines=field_lines)

    return (
        f"请从以下文档中提取指定字段：\n\n"
        f"{schema_block}\n\n"
        f"---文档内容开始---\n{text}\n---文档内容结束---\n\n"
        f"直接返回 JSON，不含任何解释或 Markdown 标记。"
    )


# ══════════════════════════════════════════════════════════════════════════════
# LLM CALL
# ══════════════════════════════════════════════════════════════════════════════

def _call_llm(
    text: str,
    fields: List[TemplateField],
    settings: ExtractionSettings,
) -> dict:
    """
    Call the LLM and return the parsed JSON dict.
    Supports OpenAI and Anthropic models via LangChain.
    Raises _ParseError on malformed JSON.
    Raises AgentError on connection/auth failures.
    """
    user_prompt = _build_user_prompt(text, fields)

    model_name = settings.model or "gpt-4o-mini"

    if model_name.startswith("gpt") or model_name.startswith("o1") or model_name.startswith("o3"):
        response_text = _call_openai(model_name, user_prompt, settings)
    elif model_name.startswith("claude"):
        response_text = _call_anthropic(model_name, user_prompt, settings)
    else:
        # Generic LangChain fallback
        response_text = _call_langchain(model_name, user_prompt, settings)

    return _parse_json_response(response_text)


def _call_openai(model: str, user_prompt: str, settings: ExtractionSettings) -> str:
    try:
        from langchain_openai import ChatOpenAI
        from langchain_core.messages import HumanMessage, SystemMessage
    except ImportError:
        raise AgentError("langchain-openai 未安装。请运行：pip install langchain-openai")

    llm = ChatOpenAI(
        model=model,
        temperature=0,
        max_tokens=settings.max_tokens or 4096,
        timeout=settings.timeout_seconds or 120,
        max_retries=2,
    )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=user_prompt),
    ]

    try:
        response = llm.invoke(messages)
        return response.content
    except Exception as e:
        raise AgentError(f"OpenAI 调用失败：{e}") from e


def _call_anthropic(model: str, user_prompt: str, settings: ExtractionSettings) -> str:
    try:
        from langchain_anthropic import ChatAnthropic
        from langchain_core.messages import HumanMessage, SystemMessage
    except ImportError:
        raise AgentError("langchain-anthropic 未安装。请运行：pip install langchain-anthropic")

    llm = ChatAnthropic(
        model=model,
        temperature=0,
        max_tokens=settings.max_tokens or 4096,
        timeout=settings.timeout_seconds or 120,
    )

    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=user_prompt),
    ]

    try:
        response = llm.invoke(messages)
        return response.content
    except Exception as e:
        raise AgentError(f"Anthropic 调用失败：{e}") from e


def _call_langchain(model: str, user_prompt: str, settings: ExtractionSettings) -> str:
    """Generic fallback using ChatOpenAI-compatible interface via custom base_url."""
    try:
        from langchain_openai import ChatOpenAI
        from langchain_core.messages import HumanMessage, SystemMessage
    except ImportError:
        raise AgentError("langchain-openai 未安装")

    llm = ChatOpenAI(
        model=model,
        temperature=0,
        max_tokens=settings.max_tokens or 4096,
    )
    messages = [
        SystemMessage(content=_SYSTEM_PROMPT),
        HumanMessage(content=user_prompt),
    ]
    try:
        response = llm.invoke(messages)
        return response.content
    except Exception as e:
        raise AgentError(f"LLM 调用失败 ({model})：{e}") from e


# ══════════════════════════════════════════════════════════════════════════════
# JSON PARSING
# ══════════════════════════════════════════════════════════════════════════════

_RE_CODE_FENCE = re.compile(r"```(?:json)?\s*([\s\S]*?)```", re.IGNORECASE)
_RE_JSON_OBJECT = re.compile(r"\{[\s\S]*\}", re.DOTALL)


def _parse_json_response(text: str) -> dict:
    """
    Extract and parse JSON from LLM output.
    Handles: code fences, leading/trailing prose, single-quoted JSON.
    Raises _ParseError if no valid JSON is found.
    """
    if not text:
        raise _ParseError("LLM 返回了空响应。")

    # Try stripping code fences first
    fence_match = _RE_CODE_FENCE.search(text)
    candidate = fence_match.group(1).strip() if fence_match else text.strip()

    # Try direct parse
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass

    # Try extracting the outermost JSON object
    obj_match = _RE_JSON_OBJECT.search(candidate)
    if obj_match:
        try:
            return json.loads(obj_match.group())
        except json.JSONDecodeError:
            pass

    # Last resort: fix single quotes (some models output Python-style dicts)
    try:
        fixed = candidate.replace("'", '"').replace("None", "null").replace("True", "true").replace("False", "false")
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    raise _ParseError(f"无法解析 LLM 输出为 JSON。前 300 字符：{text[:300]!r}")


# ══════════════════════════════════════════════════════════════════════════════
# FIELD PARSING + VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def _parse_and_validate(
    raw: dict,
    fields: List[TemplateField],
    doc: ExtractedDocument,
) -> Tuple[List[FieldValue], List[ResultIssue]]:
    """
    Convert the raw LLM JSON into typed FieldValue objects and collect issues.
    """
    values: List[FieldValue] = []
    issues: List[ResultIssue] = []

    for fdef in fields:
        # LLM returns either {"value": ..., "source_snippet": ...} or raw value
        raw_entry = raw.get(fdef.key)
        if isinstance(raw_entry, dict):
            raw_value  = raw_entry.get("value")
            snippet    = str(raw_entry.get("source_snippet") or "")[:150]
        else:
            raw_value  = raw_entry
            snippet    = ""

        # Verify snippet actually exists in document (prevent hallucination)
        if snippet and len(snippet) > 20:
            snippet = _verify_snippet(snippet, doc)

        # Type coercion
        coerced, cast_issue = _coerce_value(raw_value, fdef)
        if cast_issue:
            issues.append(ResultIssue(
                field_key=fdef.key,
                field_name=fdef.name,
                issue_type=IssueType.TYPE_INVALID,
                message=cast_issue,
                raw_value=raw_value,
            ))

        # Required check
        if fdef.required and coerced is None:
            issues.append(ResultIssue(
                field_key=fdef.key,
                field_name=fdef.name,
                issue_type=IssueType.REQUIRED_MISSING,
                message=f"必填字段「{fdef.name}」未找到对应内容。",
                raw_value=raw_value,
            ))

        values.append(FieldValue(
            key=fdef.key,
            name=fdef.name,
            value=coerced,
            raw_ai_value=raw_value,
            source_snippet=snippet or None,
            manually_edited=False,
            edited_at=None,
        ))

    return values, issues


def _coerce_value(raw: Any, fdef: TemplateField) -> Tuple[Any, Optional[str]]:
    """
    Attempt to cast the raw LLM value to the field's declared type.
    Returns (coerced_value, error_message_or_None).
    """
    if raw is None or raw == "" or raw == "null":
        return None, None

    ftype = fdef.type

    try:
        if ftype == FieldType.TEXT:
            return str(raw).strip(), None

        elif ftype == FieldType.LONG_TEXT:
            return str(raw).strip(), None

        elif ftype == FieldType.INTEGER:
            return _to_int(raw), None

        elif ftype == FieldType.DECIMAL:
            return _to_float(raw), None

        elif ftype == FieldType.DATE:
            return _to_date(raw), None

        elif ftype == FieldType.BOOLEAN:
            return _to_bool(raw), None

        elif ftype == FieldType.LIST:
            return _to_list(raw), None

        else:
            return str(raw).strip(), None

    except (ValueError, TypeError) as e:
        return None, f"类型转换失败（期望 {ftype.value}）：{e}"


# ── Type coercion helpers ─────────────────────────────────────────────────────

_RE_DIGITS = re.compile(r"[\d,，\s]+")
_RE_NUMBER = re.compile(r"[-+]?\d[\d,，]*\.?\d*")


def _to_int(raw: Any) -> int:
    if isinstance(raw, int):
        return raw
    if isinstance(raw, float):
        return int(raw)
    s = str(raw).replace(",", "").replace("，", "").replace(" ", "").strip()
    m = re.search(r"[-+]?\d+", s)
    if m:
        return int(m.group())
    raise ValueError(f"无法解析为整数：{raw!r}")


def _to_float(raw: Any) -> float:
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).replace(",", "").replace("，", "").replace(" ", "").strip()
    # Strip common currency symbols
    for sym in ("¥", "￥", "$", "€", "£", "₹", "元", "万元", "亿"):
        s = s.replace(sym, "")
    m = re.search(r"[-+]?\d+\.?\d*", s)
    if m:
        return float(m.group())
    raise ValueError(f"无法解析为数字：{raw!r}")


def _to_date(raw: Any) -> Optional[str]:
    """Normalise various date formats to YYYY-MM-DD string."""
    if raw is None:
        return None
    s = str(raw).strip()

    # Already correct
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s

    # Common patterns
    patterns = [
        (r"(\d{4})[年/\-\.](\d{1,2})[月/\-\.](\d{1,2})[日]?", "{}-{:02d}-{:02d}"),
        (r"(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})",           "{2}-{0:02d}-{1:02d}"),
        (r"(\d{4})(\d{2})(\d{2})",                               "{}-{}-{}"),
    ]

    for pattern, fmt in patterns:
        m = re.search(pattern, s)
        if m:
            parts = [int(x) for x in m.groups()]
            try:
                if "{2}" in fmt:
                    result = f"{parts[2]:04d}-{parts[0]:02d}-{parts[1]:02d}"
                else:
                    result = fmt.format(*parts)
                # Validate
                datetime.strptime(result, "%Y-%m-%d")
                return result
            except ValueError:
                continue

    # Try dateutil as last resort
    try:
        from dateutil import parser as dp
        return dp.parse(s, dayfirst=False).strftime("%Y-%m-%d")
    except Exception:
        raise ValueError(f"无法解析日期：{raw!r}")


def _to_bool(raw: Any) -> Optional[bool]:
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, int):
        return bool(raw)
    s = str(raw).strip().lower()
    if s in ("true", "yes", "是", "有", "同意", "1", "√", "✓", "对"):
        return True
    if s in ("false", "no", "否", "无", "不", "不同意", "0", "×", "✗", "错"):
        return False
    raise ValueError(f"无法解析为布尔值：{raw!r}")


def _to_list(raw: Any) -> List[str]:
    if isinstance(raw, list):
        return [str(item).strip() for item in raw if item is not None]
    if isinstance(raw, str):
        # Try JSON first
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed]
        except Exception:
            pass
        # Split on common delimiters
        for sep in ("；", ";", "，", ",", "\n", "、"):
            if sep in raw:
                return [x.strip() for x in raw.split(sep) if x.strip()]
        return [raw.strip()]
    return [str(raw).strip()]


# ══════════════════════════════════════════════════════════════════════════════
# SNIPPET VERIFICATION (anti-hallucination)
# ══════════════════════════════════════════════════════════════════════════════

def _verify_snippet(snippet: str, doc: ExtractedDocument) -> str:
    """
    Check that the snippet actually appears in the document text.
    If not found verbatim, try to find the closest real fragment.
    Returns the verified snippet, or empty string if nothing matches.
    """
    full = doc.full_text

    # Verbatim check (strip whitespace differences)
    clean_snippet = " ".join(snippet.split())
    clean_full    = " ".join(full.split())

    if clean_snippet in clean_full:
        return snippet

    # Fuzzy: check a 15-char core substring
    core = clean_snippet[len(clean_snippet)//4: len(clean_snippet)*3//4]
    if core and core in clean_full:
        # Return the actual surrounding text from the document
        idx = clean_full.find(core)
        start = max(0, idx - 60)
        end   = min(len(clean_full), idx + len(core) + 60)
        return clean_full[start:end]

    # Could not verify — return empty to avoid surfacing hallucinated snippets
    return ""


# ══════════════════════════════════════════════════════════════════════════════
# CUSTOM EXCEPTIONS
# ══════════════════════════════════════════════════════════════════════════════

class AgentError(RuntimeError):
    """Raised when extraction cannot be completed after all retries."""


class _ParseError(ValueError):
    """Internal: raised when the LLM output cannot be parsed as JSON."""


# ══════════════════════════════════════════════════════════════════════════════
# STANDALONE TEST HELPER
# ══════════════════════════════════════════════════════════════════════════════

def _test_agent(pdf_path: str, template_xlsx: str):
    """
    Quick CLI smoke test.
    Usage:
        python agent.py contract.pdf template.xlsx
    Requires OPENAI_API_KEY env var.
    """
    import os
    from extractor import extract
    from template_service import parse_excel, save_template

    print(f"[1] Extracting text from {pdf_path}…")
    with open(pdf_path, "rb") as fh:
        raw = fh.read()
    doc = extract(pdf_path, raw)
    print(f"    {doc.page_count} pages, {doc.total_chars:,} chars")

    print(f"[2] Parsing template {template_xlsx}…")
    with open(template_xlsx, "rb") as fh:
        xraw = fh.read()
    parse_resp = parse_excel(xraw, template_xlsx)
    snap = save_template(parse_resp.upload_id, parse_resp.fields, name="test")
    print(f"    {len(snap.fields)} fields")

    print("[3] Running AI extraction…")
    settings = ExtractionSettings(model=os.getenv("AGENT_MODEL", "gpt-4o-mini"))
    result = run_extraction(doc, snap.fields, settings, file_id="test_001")

    print(f"\n✅ Done in {result.elapsed_seconds}s  (retries: {result.retry_count})")
    print(f"   Issues: {result.issue_count}")
    print("\n── Extracted Fields ─────────────────────────────────────────────")
    for fv in result.fields:
        edited_flag = " ✎" if fv.manually_edited else ""
        print(f"  {fv.name:<30} {str(fv.value):<40}{edited_flag}")
        if fv.source_snippet:
            print(f"  {'':30} └─ 来源: {fv.source_snippet[:60]}…")

    if result.issues:
        print("\n── Issues ────────────────────────────────────────────────────")
        for issue in result.issues:
            print(f"  ⚠ [{issue.field_name}] {issue.issue_type.value}: {issue.message}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Usage: python agent.py <document.pdf> <template.xlsx>")
        sys.exit(1)
    _test_agent(sys.argv[1], sys.argv[2])