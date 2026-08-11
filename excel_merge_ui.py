# -*- coding: utf-8 -*-
"""Streamlit page for local multi-file Excel merging."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from excel_merge import (
    _clean_dataframe_for_merge,
    merge_excel_dataframes,
    preview_merge,
    read_excel_sheet_dataframe,
    read_excel_sheet_names,
    to_excel_bytes,
)


def page_excel_merge() -> None:
    st.markdown("## 📊 Excel 合并")
    st.caption("上传多个结构相近的 Excel 文件，自动识别表头、对齐字段并导出汇总结果。")
    st.info("字段匹配仅使用本地规则，不调用云端 API，也不会影响本机 Ollama。")

    uploaded_files = st.file_uploader(
        "上传 Excel 文件",
        type=["xlsx"],
        accept_multiple_files=True,
        key="excel_merge_uploads",
        help="至少选择两个 .xlsx 文件；可为每个文件选择一个工作表。",
    )

    if not uploaded_files:
        st.markdown(
            "支持自动处理：表头不在第一行、空白列、重复表头行、旧“序号”列，以及字段顺序不同的表格。"
        )
        return

    selected_files = []
    for index, uploaded in enumerate(uploaded_files):
        file_bytes = uploaded.getvalue()
        try:
            sheet_names = read_excel_sheet_names(file_bytes, uploaded.name)
        except Exception as exc:
            st.error(f"无法读取 {uploaded.name}：{exc}")
            continue

        selected_sheet = st.selectbox(
            f"{uploaded.name}：选择工作表",
            sheet_names,
            key=f"excel_merge_sheet_{index}",
        )
        selected_files.append((uploaded.name, file_bytes, selected_sheet))

    if len(selected_files) < 2:
        st.warning("请至少成功读取并选择两个 Excel 文件后再合并。")
        return

    current_signature = tuple(
        (filename, len(file_bytes), sheet_name)
        for filename, file_bytes, sheet_name in selected_files
    )
    if st.session_state.get("excel_merge_signature") != current_signature:
        st.session_state.pop("excel_merge_result", None)
        st.session_state["excel_merge_signature"] = current_signature

    previews, errors, canonical_headers, similarity, exact_match = preview_merge(selected_files)
    if errors:
        for error in errors:
            st.error(error)
        return

    col1, col2, col3 = st.columns(3)
    col1.metric("参与合并", f"{len(previews)} 个工作表")
    col2.metric("表头相似度", f"{similarity:.0%}")
    col3.metric("字段匹配", "完全一致" if exact_match else "已自动对齐")

    st.markdown("### 字段映射预览")
    mapping_rows = []
    for preview in previews:
        for source_column, target_column in preview.mapping.items():
            mapping_rows.append(
                {
                    "文件": preview.filename,
                    "工作表": preview.sheet_name,
                    "原字段": source_column,
                    "合并后字段": target_column,
                }
            )
    if mapping_rows:
        st.dataframe(
            pd.DataFrame(mapping_rows),
            use_container_width=True,
            hide_index=True,
        )

    if not exact_match:
        st.warning(
            "部分字段名称或顺序不同。无法确认的字段会保留为独立列，避免丢失数据。"
        )

    drop_duplicates = st.checkbox(
        "删除完全重复的数据行",
        value=False,
        key="excel_merge_deduplicate",
    )

    if st.button("开始合并并生成 Excel", type="primary", use_container_width=True):
        try:
            dataframes = []
            mappings = []

            for preview, (filename, file_bytes, sheet_name) in zip(previews, selected_files):
                dataframe = read_excel_sheet_dataframe(file_bytes, sheet_name, filename)
                dataframe = _clean_dataframe_for_merge(dataframe)
                if dataframe.empty:
                    continue

                dataframe["来源文件"] = filename
                dataframes.append(dataframe)
                mappings.append(preview.mapping)

            if not dataframes:
                st.error("所选工作表均没有可合并的数据行。")
                return

            merged = merge_excel_dataframes(
                dataframes=dataframes,
                mappings=mappings,
                canonical_headers=canonical_headers,
                drop_duplicates=drop_duplicates,
            )

            if "序号" in merged.columns:
                merged = merged.drop(columns=["序号"])
            merged.insert(0, "序号", range(1, len(merged) + 1))

            st.session_state["excel_merge_result"] = merged
            st.success(f"合并完成：{len(dataframes)} 个工作表，生成 {len(merged)} 行数据。")
        except Exception as exc:
            st.exception(exc)

    merged_result = st.session_state.get("excel_merge_result")
    if isinstance(merged_result, pd.DataFrame):
        st.markdown("### 合并结果预览")
        st.dataframe(
            merged_result,
            use_container_width=True,
            hide_index=True,
            height=420,
        )

        st.download_button(
            "⬇️ 下载合并结果 Excel",
            data=to_excel_bytes(merged_result),
            file_name="Excel合并结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
