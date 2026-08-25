"""专属公文格式转换 Skill：将结构化原稿渲染为标准公文 DOCX。"""
from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any

from ai_writer import TEMP_DIR
from converter import build_structure, convert_with_roles


def format_structured_documents(documents: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    """唯一格式化入口，复用平台公文格式转换规则与预览结构。"""
    formatted: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    for document in documents:
        try:
            raw_path = Path(document.get("rawPath") or document["path"])
            if not raw_path.exists():
                raise RuntimeError("原始公文不存在")
            result_id = uuid.uuid4().hex
            result_path = TEMP_DIR / f"{result_id}_ai.docx"
            roles = document.get("roles") or {}
            conversion = convert_with_roles(
                str(raw_path), str(result_path), roles=roles,
                template_id=document.get("templateId"), split_embedded_markers=False,
                split_inline_headings=False, force_heading_bold=True,
            )
            if not result_path.exists():
                raise RuntimeError("专属格式转换工具未生成成品文件")
            paragraphs = build_structure(str(result_path), document.get("templateId"), split_embedded_markers=False)
            for paragraph in paragraphs:
                role = roles.get(paragraph["index"]) or roles.get(str(paragraph["index"]))
                if role:
                    paragraph.update({"role": role, "confidence": 1.0, "lowConfidence": False})
            document.update({
                "documentId": result_id, "path": str(result_path), "paragraphs": paragraphs,
                "formatted": True, "formatting": {"tool": "convert_with_roles", "log": conversion.get("log", [])},
            })
            formatted.append(document)
        except Exception as exc:
            failures.append({"filename": document.get("sourceName") or document.get("filename") or "未知材料", "reason": f"格式转换失败：{exc}"})
    if not formatted:
        raise RuntimeError("全部公文格式转换失败。" + "；".join(item["reason"] for item in failures[:3]))
    return formatted, failures
