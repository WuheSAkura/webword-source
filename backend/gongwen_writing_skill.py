"""受限公文写作 Skill：模板解析、意图约束与结构化写作的唯一入口。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from ai_writer import (
    MAX_REFERENCE_INPUT_CHARS,
    extract_document_inventory,
    generate_documents,
    get_template_by_key,
    inventory_to_text,
    load_template_references,
)

MAX_TEMPLATE_RULE_CHARS = MAX_REFERENCE_INPUT_CHARS


def load_preparsed_template_rules(template_id: str, template_key: str | None = None) -> dict[str, Any]:
    """从模板库读取后台入库时已完成的结构化解析结果。"""
    template = get_template_by_key(template_key, template_id)
    if not template:
        raise ValueError(f"未找到模板分类：{template_id}")
    references = [item for item in load_template_references(template["id"], template["sourceDir"])
                  if item.get("status") == "ready" and item.get("structureJson")]
    if not references:
        raise ValueError(f"模板分类“{template.get('label') or template_id}”尚无预解析范文")
    return {
        "templates": [{"filename": item["filename"], "text": item.get("text", ""),
                       "structureJson": item.get("structureJson") or {}, "roles": item.get("roles") or []}
                      for item in references],
        "warnings": [item["warning"] for item in references if item.get("warning")],
        "count": len(references),
        "source": "preparsed_template_library",
        "contextChars": sum(len(item.get("text", "")) for item in references),
    }


def parse_reference_templates(paths: list[Path]) -> dict[str, Any]:
    """逐份解析临时模板，合并为本次任务的规则上下文。"""
    rules: list[dict[str, Any]] = []
    warnings: list[str] = []
    remaining_chars = MAX_TEMPLATE_RULE_CHARS
    for path in paths:
        inventory = extract_document_inventory(path)
        text = inventory_to_text(inventory).strip()
        if not text:
            warnings.append(f"模板 {path.name} 未提取到可用正文")
            continue
        if remaining_chars <= 0:
            warnings.append(f"模板 {path.name} 未纳入模型上下文：已达到模板规则字符上限")
            continue
        included_text = text[:remaining_chars]
        if len(included_text) < len(text):
            warnings.append(f"模板 {path.name} 已按规则上下文上限截取")
        rules.append({
            "filename": path.name,
            "text": included_text,
            "inventory": inventory,
        })
        remaining_chars -= len(included_text)
        warnings.extend(inventory.get("warnings") or [])
    if paths and not rules:
        raise ValueError("上传的模板均无法解析")
    return {"templates": rules, "warnings": warnings, "count": len(rules),
            "contextChars": sum(len(item["text"]) for item in rules)}


def run_writing_skill(*, prompt: str, material_paths: list[Path], template_paths: list[Path],
                      request_url: str, api_key: str, model_name: str,
                      template_id: str | None, template_key: str | None = None,
                      temperature: float, speed_mode: str,
                      strict_reference_isolation: bool = False,
                      allow_degradation: bool = False,
                      template_rules: dict[str, Any] | None = None,
                      text_materials: list[dict[str, str]] | None = None) -> dict[str, Any]:
    """由调度器调用的唯一写作入口。"""
    template_rules = template_rules if template_rules is not None else parse_reference_templates(template_paths)
    # 用户 prompt 只保留写作要求，不把范文结构/原文痕迹拼进去。
    # 体例坐标由 generate_documents → select_reference → reference_structure_payload 注入，
    # 避免材料润色/范文匹配把模板文件内容误当成上传材料。
    user_prompt = (prompt or "根据材料生成正式公文").strip()
    result = generate_documents(
        prompt=user_prompt,
        upload_paths=material_paths,
        text_materials=text_materials or [],
        request_url=request_url,
        api_key=api_key,
        model_name=model_name,
        template_id=template_id,
        template_key=template_key,
        temperature=temperature,
        speed_mode=speed_mode,
        strict_reference_isolation=strict_reference_isolation,
        allow_degradation=allow_degradation,
        format_output=False,
    )
    result["templateRules"] = {
        "count": template_rules["count"],
        "filenames": [item["filename"] for item in template_rules["templates"]],
        "warnings": template_rules["warnings"],
        "contextChars": template_rules.get("contextChars", 0),
    }
    return result
