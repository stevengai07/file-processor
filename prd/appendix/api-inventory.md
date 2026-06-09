# API 清单

## 现有接口

| 方法 | 路径 | 状态 | 说明 |
|---|---|---|---|
| POST | `/extract` | 已实现 | 单文件 PDF/DOCX 提取，固定返回 `ExtractedInfo` |
| GET | `/health` | 已实现 | 返回服务状态和 OCR 预设 |

现有 `/extract` 未接收前端选择的模型，内部使用默认模型；不支持动态字段、批量任务、进度、编辑或汇总导出。

## MVP 建议接口

### `POST /api/templates/parse`

请求：`multipart/form-data`，包含 Excel 文件、可选工作表和表头行。

响应：

```json
{
  "upload_id": "tmp_123",
  "sheets": ["字段定义"],
  "selected_sheet": "字段定义",
  "fields": [
    {
      "key": "contract_number",
      "name": "合同编号",
      "description": "合同的唯一编号",
      "type": "text",
      "required": true,
      "example": "HT-2026-001",
      "prompt_hint": "可能写作合同号、协议编号",
      "order": 1
    }
  ],
  "errors": []
}
```

### `POST /api/templates`

保存用户确认后的字段数组，返回 `template_id`、版本和创建时间。

### `POST /api/tasks`

请求包含 `template_id`、任务名称、模型、OCR 预设、OCR 语言和并发数。返回 `task_id` 和任务状态。

### `POST /api/tasks/{task_id}/files`

接收多个 `files`。响应必须逐文件返回接受或拒绝结果，不能因一个非法文件使整个请求失败。

### `POST /api/tasks/{task_id}/start`

将任务从 `draft` 转为 `queued/running`。重复调用应具备幂等性。

### `GET /api/tasks/{task_id}`

返回任务状态、总数、等待数、处理中、成功、待检查、失败、已取消和百分比。

### `GET /api/tasks/{task_id}/results`

查询参数：`page`、`page_size`、`status`、`keyword`、`has_issues`。

每条结果包含文件元数据、动态字段、问题列表、人工修改标记和更新时间。

### `PATCH /api/tasks/{task_id}/results/{file_id}`

请求仅包含需要修改的字段。后端根据模板校验字段存在性和类型，并重新计算文件状态。

### `POST /api/tasks/{task_id}/exports/{format}`

`format` 为 `excel` 或 `docx`。请求包含导出范围和文件 ID；成功时返回文件流或可下载地址。

## 统一错误结构

```json
{
  "code": "INVALID_TEMPLATE",
  "message": "字段名称不能为空",
  "details": {
    "row": 4,
    "field": "name"
  }
}
```

## 技术约束

- 批量处理不应在单个 HTTP 请求中同步完成。
- 生产环境建议使用任务队列和持久化数据库。
- 动态字段不能直接依赖当前固定 `ExtractedInfo` 模型，应使用模板驱动结构校验。
- 上传文件名必须清洗，临时路径不得由用户输入直接拼接。

