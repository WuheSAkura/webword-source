"""轻量公文流程调度器：严格串行推进并记录节点快照。"""
from __future__ import annotations

import uuid
from time import perf_counter
from pathlib import Path
from typing import Any, Callable

from gongwen_formatting_skill import format_structured_documents
from gongwen_writing_skill import load_preparsed_template_rules, parse_reference_templates, run_writing_skill

NODE_ORDER = ("uploaded", "template_parsed", "writing_skill_completed", "formatted", "exported")
TASKS: dict[str, dict[str, Any]] = {}


def _task(task_id: str) -> dict[str, Any]:
    return TASKS[task_id]


def _node(task: dict[str, Any], name: str, status: str, **extra: Any) -> None:
    item = task["nodes"].setdefault(name, {"name": name})
    item.update({"status": status, **extra})
    if status == "running":
        item["startedAt"] = perf_counter()
    elif status in {"completed", "failed", "blocked"} and item.get("startedAt") is not None:
        item["elapsedSeconds"] = round(perf_counter() - item["startedAt"], 2)


def run_workflow(*, prompt: str, material_paths: list[Path], template_paths: list[Path],
                 request_url: str, api_key: str, model_name: str, template_id: str | None,
                 temperature: float, speed_mode: str, strict_reference_isolation: bool = False,
                 export_documents: Callable[[list[dict[str, Any]]], list[dict[str, Any]]] | None = None) -> dict[str, Any]:
    task_id = uuid.uuid4().hex
    task = {"taskId": task_id, "status": "running", "nodes": {}, "documents": [], "failures": []}
    TASKS[task_id] = task
    started = perf_counter()
    try:
        _node(task, "uploaded", "running", materialCount=len(material_paths), templateCount=len(template_paths))
        if not material_paths:
            raise ValueError("请至少上传一份业务材料后再生成")
        _node(task, "uploaded", "completed", materialCount=len(material_paths), templateCount=len(template_paths))
        _node(task, "template_parsed", "running")
        template_rules = load_preparsed_template_rules(template_id) if template_id else parse_reference_templates(template_paths)
        _node(task, "template_parsed", "completed", summary={
            "count": template_rules["count"],
            "filenames": [item["filename"] for item in template_rules["templates"]],
            "warnings": template_rules["warnings"],
        })
        _node(task, "writing_skill_completed", "running")
        result = run_writing_skill(
            prompt=prompt, material_paths=material_paths, template_paths=template_paths,
            request_url=request_url, api_key=api_key, model_name=model_name,
            template_id=template_id, temperature=temperature, speed_mode=speed_mode,
            strict_reference_isolation=strict_reference_isolation,
            template_rules=template_rules,
        )
        _node(task, "writing_skill_completed", "completed", documentCount=len(result.get("documents", [])))
        documents = result.get("documents", [result])
        _node(task, "formatted", "running", inputCount=len(documents), tool="gongwen_formatting_skill")
        documents, formatting_failures = format_structured_documents(documents)
        _node(task, "formatted", "completed", ruleSource="convert_with_roles", documentCount=len(documents),
              previewReady=all(bool(item.get("paragraphs")) for item in documents))
        task["documents"] = documents
        task["failures"] = [*(result.get("failures") or []), *formatting_failures]
        result["failures"] = task["failures"]
        _node(task, "exported", "running", documentCount=len(documents))
        if export_documents:
            documents = export_documents(documents)
            result["documents"] = documents
            result.update(documents[0])
        if not all(Path(item.get("path", "")).exists() and item.get("fileId") for item in documents):
            raise RuntimeError("Word 导出登记失败，未生成格式转换工作台下载标识")
        task["documents"] = documents
        _node(task, "exported", "completed", documentCount=len(documents),
              downloadable=True)
        task.update({"status": "completed", "elapsedSeconds": round(perf_counter() - started, 2)})
        for node in task["nodes"].values():
            node.pop("startedAt", None)
        return {"success": True, **task, **result}
    except Exception as exc:
        current = next((name for name in NODE_ORDER if task["nodes"].get(name, {}).get("status") == "running"), None)
        if current:
            _node(task, current, "failed", error=str(exc))
            index = NODE_ORDER.index(current)
            for name in NODE_ORDER[index + 1:]:
                _node(task, name, "blocked", error=f"前置节点 {current} 失败")
        task.update({"status": "failed", "error": str(exc), "elapsedSeconds": round(perf_counter() - started, 2)})
        for node in task["nodes"].values():
            node.pop("startedAt", None)
        return {"success": False, **task}


def get_workflow_task(task_id: str) -> dict[str, Any] | None:
    task = TASKS.get(task_id)
    return dict(task) if task else None


def attach_export_records(task_id: str, documents: list[dict[str, Any]]) -> None:
    """将格式转换工作台登记后的 fileId 回传到调度任务。"""
    task = TASKS.get(task_id)
    if not task:
        return
    task["documents"] = documents
    task["exportRecords"] = [
        {
            "documentId": item.get("documentId"),
            "fileId": item.get("fileId"),
            "filename": item.get("filename"),
            "preview": item.get("paragraphs") or [],
        }
        for item in documents
    ]
    task["nodes"].setdefault("formatted", {})["previewData"] = task["exportRecords"]
    task["nodes"].setdefault("exported", {})["downloadRecords"] = task["exportRecords"]
