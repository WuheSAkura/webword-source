"""FastAPI 后端入口 —— Word 公文格式转换服务

流程：上传 → 识别结构(/structure) → 人工校正 → 按校正套格式(/convert) → 预览/下载
"""

import json
import re
import uuid
import shutil
import subprocess
from contextlib import asynccontextmanager
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from docx import Document
from docx.oxml.ns import qn

from config import (
    ROLE_LABELS,
    get_default_template_id,
    get_resolved_styles,
    get_template_catalog,
    has_document_template,
    normalize_template_id,
)
from converter import (
    convert_with_roles,
    build_structure,
    classify_one,
    refine_structure,
    split_paragraph_text,
    apply_local_edit,
)
from ai_writer import (
    MAX_AI_UPLOAD_SIZE,
    REFERENCE_SUFFIXES,
    SUPPORTED_INPUT_SUFFIXES,
    build_upload_temp_path,
    generate_documents,
    get_template_by_id,
    guess_template_id,
    get_model_config,
    public_template_view,
    safe_extract_text,
    save_model_config,
    scan_template_library,
    sync_template_library,
)
from gongwen_workflow import attach_export_records, get_workflow_task, run_workflow

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
STATIC_DIR = BASE_DIR / "static"
MAX_FILE_SIZE = 20 * 1024 * 1024
MAX_UPLOAD_BATCH = 20
RENDER_TIMEOUT_SECONDS = 60
AI_DOCUMENT_REGISTRY_PATH = TEMP_DIR / "ai_document_registry.json"

# { file_id: {path, name, size, processed, result_id, result_path,
#              roles, template_id, log, undo_stack} }
uploaded_files: dict[str, dict] = {}
generated_ai_documents: dict[str, dict] = {}


class ConvertRequest(BaseModel):
    roles: dict[str, str] | None = None  # {逻辑段索引(str): 角色}
    template_id: str | None = None


class EditRequest(BaseModel):
    selection: dict
    font: dict | None = None
    paragraph: dict | None = None


class AiModelConfigRequest(BaseModel):
    request_url: str = ""
    api_key: str = ""
    model_name: str = ""
    models: list[str] | None = None


def cleanup_old_files(max_age_minutes: int = 30):
    now = datetime.now().timestamp()
    for f in TEMP_DIR.glob("*"):
        if f.name == AI_DOCUMENT_REGISTRY_PATH.name:
            continue
        if f.is_file() and (now - f.stat().st_mtime) > max_age_minutes * 60:
            try:
                f.unlink()
            except OSError:
                pass
    expired_ids = []
    for key, info in list(uploaded_files.items()):
        path = Path(info.get("path") or "")
        result_path = Path(info.get("result_path") or "")
        path_missing = bool(info.get("path")) and not path.exists()
        result_missing = info.get("processed") and info.get("result_path") and not result_path.exists()
        if path_missing or result_missing:
            expired_ids.append(key)
    for key in expired_ids:
        info = uploaded_files.pop(key, None)
        if not info:
            continue
        result_id = info.get("result_id")
        if result_id:
            generated_ai_documents.pop(result_id, None)
    stale_docs = [
        document_id
        for document_id, document in list(generated_ai_documents.items())
        if not Path(document.get("path") or "").exists()
    ]
    for document_id in stale_docs:
        generated_ai_documents.pop(document_id, None)
    if expired_ids or stale_docs:
        _persist_ai_document_registry()


def _purge_file_record(file_id: str) -> bool:
    info = uploaded_files.pop(file_id, None)
    if not info:
        return False
    result_id = info.get("result_id")
    if result_id:
        generated_ai_documents.pop(result_id, None)
    _delete_file(info.get("path"))
    _clear_result_artifacts(info)
    _persist_ai_document_registry()
    return True


def _load_ai_document_registry() -> None:
    """进程重载后恢复仍在临时目录中的 AI 生成文件登记。"""
    if not AI_DOCUMENT_REGISTRY_PATH.exists():
        return
    try:
        records = json.loads(AI_DOCUMENT_REGISTRY_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return
    valid_records = []
    for record in records if isinstance(records, list) else []:
        file_id = str(record.get("fileId") or "")
        document_id = str(record.get("documentId") or "")
        source_path = Path(str(record.get("sourcePath") or ""))
        result_path = Path(str(record.get("resultPath") or ""))
        if not re.fullmatch(r"[a-f0-9]{32}", file_id) or not document_id:
            continue
        if not source_path.exists() or not result_path.exists():
            continue
        info = {
            "path": source_path,
            "name": record.get("name") or "AI公文.docx",
            "size": source_path.stat().st_size,
            "processed": True,
            "result_id": document_id,
            "result_path": str(result_path),
            "roles": record.get("roles") or {},
            "template_id": normalize_template_id(record.get("templateId")),
            "log": ["AI 生成结果已从本地登记恢复"],
            "undo_stack": [],
        }
        uploaded_files[file_id] = info
        generated_ai_documents[document_id] = {
            "documentId": document_id,
            "path": str(result_path),
            "filename": info["name"],
            "fileId": file_id,
        }
        valid_records.append(record)
    try:
        AI_DOCUMENT_REGISTRY_PATH.write_text(json.dumps(valid_records, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _persist_ai_document_registry() -> None:
    records = []
    for document_id, document in generated_ai_documents.items():
        file_id = str(document.get("fileId") or "")
        info = uploaded_files.get(file_id)
        if not info:
            continue
        records.append({
            "documentId": document_id,
            "fileId": file_id,
            "sourcePath": str(info["path"]),
            "resultPath": str(info["result_path"]),
            "name": info["name"],
            "templateId": info.get("template_id"),
            "roles": info.get("roles") or {},
        })
    try:
        AI_DOCUMENT_REGISTRY_PATH.write_text(json.dumps(records, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _get_uploaded_file(file_id: str) -> dict | None:
    info = uploaded_files.get(file_id)
    if info and Path(info.get("path", "")).exists():
        return info
    # 清单异常或刚完成重载时，按固定 AI 源文件名进行无状态恢复。
    if not re.fullmatch(r"[a-f0-9]{32}", file_id):
        return None
    source_path = TEMP_DIR / f"{file_id}_ai_source.docx"
    if not source_path.exists():
        return None
    info = {
        "path": source_path,
        "name": "AI公文.docx",
        "size": source_path.stat().st_size,
        "processed": True,
        "result_id": None,
        "result_path": str(source_path),
        "roles": {},
        "template_id": get_default_template_id(),
        "log": ["AI 生成结果已按源文件自动恢复"],
        "undo_stack": [],
    }
    uploaded_files[file_id] = info
    return info


@asynccontextmanager
async def lifespan(app: FastAPI):
    cleanup_old_files()
    _load_ai_document_registry()
    templates = scan_template_library()
    if not templates:
        print("[warn] AI 模板库为空，请确认范文目录已复制并执行 /api/ai/templates/sync")
    elif len(templates) != 15:
        print(f"[warn] AI 模板库条目数为 {len(templates)}，期望 15；请检查 Docker 是否带入 Windows 路径 db 或缺少范文目录")
    yield


app = FastAPI(title="Word公文格式转换工具", version="2.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _validate_template_id(template_id: str | None) -> str:
    if template_id is None:
        return get_default_template_id()
    if not has_document_template(template_id):
        raise HTTPException(400, detail="未知的公文模板")
    return normalize_template_id(template_id)


def _delete_file(path_value) -> None:
    if not path_value:
        return
    try:
        path = Path(path_value)
        if path.exists():
            path.unlink()
    except OSError:
        pass


def _clear_result_artifacts(info: dict) -> None:
    result_path = info.get("result_path")
    if result_path and str(result_path) != str(info.get("path")):
        _delete_file(result_path)
    _delete_file(info.get("render_pdf_path"))
    for undo_path in info.get("undo_stack") or []:
        _delete_file(undo_path)
    info["undo_stack"] = []
    info["render_pdf_path"] = None


def _register_ai_document_for_converter(document: dict) -> dict:
    result_path = Path(document["path"])
    file_id = uuid.uuid4().hex
    source_path = TEMP_DIR / f"{file_id}_ai_source.docx"
    shutil.copyfile(result_path, source_path)
    file_size = result_path.stat().st_size if result_path.exists() else source_path.stat().st_size
    template_id = normalize_template_id(document.get("templateId"))
    roles = {
        str(item.get("index")): item.get("role")
        for item in document.get("paragraphs", [])
        if item.get("index") is not None and item.get("role")
    }
    uploaded_files[file_id] = {
        "path": source_path,
        "name": document.get("filename") or "AI公文.docx",
        "size": file_size,
        "processed": True,
        "result_id": document["documentId"],
        "result_path": str(result_path),
        "roles": roles,
        "template_id": template_id,
        "log": ["AI 生成结果已自动同步到公文格式转换工具"],
        "undo_stack": [],
    }
    return {
        "id": file_id,
        "name": uploaded_files[file_id]["name"],
        "size": file_size,
        "templateId": template_id,
        "documentId": document["documentId"],
    }


def build_result_preview(file_path: str, roles: dict | None = None,
                         template_id: str | None = None) -> list[dict]:
    """从套好格式的 docx 提取段落预览（含 run 字体），供结果预览与局部编辑。"""
    template_id = normalize_template_id(template_id)
    corrected_roles = {int(k): v for k, v in (roles or {}).items()}
    doc = Document(str(file_path))
    logical_items = []
    for source_index, paragraph in enumerate(doc.paragraphs):
        raw = paragraph.text.strip()
        if not raw:
            continue
        for part in split_paragraph_text(raw):
            if part.strip():
                logical_items.append((source_index, paragraph, part.strip()))

    total = len(logical_items)

    def extract_runs(paragraph, text):
        if paragraph.text.strip() != text:
            return [{"text": text, "fontEast": "", "fontWest": "", "size": None, "bold": None}]
        runs = []
        for run in paragraph.runs:
            if not run.text:
                continue
            east = ""
            rPr = run._element.rPr
            if rPr is not None:
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is not None:
                    east = rFonts.get(qn("w:eastAsia")) or ""
            runs.append({
                "text": run.text,
                "fontEast": east,
                "fontWest": run.font.name or "",
                "size": run.font.size.pt if run.font.size else None,
                "bold": run.font.bold,
            })
        return runs or [{"text": text, "fontEast": "", "fontWest": "", "size": None, "bold": None}]

    def xml_number(element, attr_name: str, divisor: float = 1.0):
        if element is None:
            return None
        raw = element.get(qn(attr_name))
        try:
            return round(float(raw) / divisor, 2) if raw is not None else None
        except (TypeError, ValueError):
            return None

    def extract_layout(paragraph):
        align_map = {None: "left", 0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
        layout = {
            "align": align_map.get(paragraph.alignment, "left"),
            "lineRule": None,
            "linePt": None,
            "lineMultiple": None,
            "firstLineChars": 0,
            "leftChars": 0,
            "rightChars": 0,
        }
        pPr = paragraph._p.find(qn("w:pPr"))
        if pPr is None:
            return layout
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            line_rule = spacing.get(qn("w:lineRule")) or "auto"
            line_value = xml_number(spacing, "w:line")
            layout["lineRule"] = "exact" if line_rule == "exact" else ("atLeast" if line_rule == "atLeast" else "multiple")
            if line_rule in ("exact", "atLeast") and line_value is not None:
                layout["linePt"] = round(line_value / 20, 2)
            elif line_value is not None:
                layout["lineMultiple"] = round(line_value / 240, 2)
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            layout["firstLineChars"] = xml_number(ind, "w:firstLineChars", 100) or 0
            layout["leftChars"] = xml_number(ind, "w:leftChars", 100) or 0
            layout["rightChars"] = xml_number(ind, "w:rightChars", 100) or 0
        return layout

    items = []
    for i, (source_index, paragraph, text) in enumerate(logical_items):
        role, confidence = classify_one(text, i, total, template_id=template_id)
        items.append({
            "index": i,
            "text": text[:200] + ("..." if len(text) > 200 else ""),
            "fullText": text,
            "role": role,
            "sourceIndex": source_index,
            "confidence": confidence,
            "runs": extract_runs(paragraph, text),
            "layout": extract_layout(paragraph),
        })
    refine_structure(items, template_id)
    for item in items:
        role = corrected_roles.get(item["index"], item["role"])
        item["role"] = role
        item["roleLabel"] = ROLE_LABELS.get(role, "正文")
        item.pop("confidence", None)
    return items


def _find_soffice() -> str | None:
    for command in ("soffice", "libreoffice"):
        found = shutil.which(command)
        if found:
            return found
    common_paths = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for path in common_paths:
        if path.exists():
            return str(path)
    return None


def render_docx_to_pdf(docx_path: Path) -> dict:
    soffice = _find_soffice()
    if not soffice:
        return {
            "status": "unavailable",
            "message": "未找到 LibreOffice/soffice，无法进行真实分页渲染校验",
        }

    pdf_path = docx_path.with_suffix(".pdf")
    if pdf_path.exists() and pdf_path.stat().st_mtime >= docx_path.stat().st_mtime:
        return {"status": "ok", "pdfPath": str(pdf_path), "cached": True}

    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(docx_path.parent),
        str(docx_path),
    ]
    try:
        completed = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=RENDER_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return {"status": "failed", "message": f"真实渲染超时（{RENDER_TIMEOUT_SECONDS}秒）"}
    except OSError as exc:
        return {"status": "failed", "message": f"启动渲染器失败：{exc}"}

    if completed.returncode != 0 or not pdf_path.exists():
        detail = (completed.stderr or completed.stdout or "").strip()
        return {
            "status": "failed",
            "message": "真实渲染失败",
            "detail": detail[-1000:],
        }

    return {"status": "ok", "pdfPath": str(pdf_path), "cached": False}


# ── API ──
@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.get("/api/config")
async def get_format_config():
    """返回格式配置与角色标签，供前端校正/编辑面板使用（消除前端硬编码）。"""
    default_template_id = get_default_template_id()
    return {
        "success": True,
        "styles": get_resolved_styles(default_template_id),
        "labels": ROLE_LABELS,
        "defaultTemplateId": default_template_id,
        "templates": get_template_catalog(),
    }


@app.get("/api/ai/templates")
async def get_ai_templates():
    """返回 AI 写作可参考的模板库索引。"""
    return {
        "success": True,
        "templates": [public_template_view(item) for item in scan_template_library()],
    }


@app.post("/api/ai/templates/sync")
async def sync_ai_templates(enrich: bool = False):
    templates = sync_template_library(enrich_structures=enrich)
    return {
        "success": True,
        "count": len(templates),
        "templates": [public_template_view(item) for item in templates],
    }


def _unique_target_path(directory: Path, filename: str) -> Path:
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    candidate = directory / filename
    index = 1
    while candidate.exists():
        candidate = directory / f"{stem}_{index}{suffix}"
        index += 1
    return candidate


@app.post("/api/ai/templates/upload")
async def upload_ai_templates(
    files: list[UploadFile] = File(...),
    template_id: str = Form(default=""),
):
    saved = []
    skipped = []
    assignments = []
    for file in files:
        filename = Path(file.filename or "").name
        suffix = Path(filename).suffix.lower()
        if not filename or filename.startswith("~$") or suffix not in REFERENCE_SUFFIXES:
            skipped.append({"filename": filename or "(unknown)", "reason": "不支持的模板文件类型"})
            continue
        content = await file.read()
        if len(content) > MAX_AI_UPLOAD_SIZE:
            raise HTTPException(400, detail=f"{filename} 超过 30MB")

        probe_path = None
        probe_text = ""
        if not template_id and suffix in SUPPORTED_INPUT_SUFFIXES:
            probe_path = TEMP_DIR / f"{uuid.uuid4().hex}{suffix}"
            with open(probe_path, "wb") as f:
                f.write(content)
            probe_text, _ = safe_extract_text(probe_path, limit=6000)
        resolved_template_id = template_id or guess_template_id(probe_text or filename, filename)
        if probe_path:
            _delete_file(probe_path)
        if resolved_template_id == "generic":
            skipped.append({"filename": filename, "reason": "自动判断无法识别文种，请选择具体模板后上传"})
            continue
        template = get_template_by_id(resolved_template_id)
        if not template:
            skipped.append({"filename": filename, "reason": f"未找到模板类别 {resolved_template_id}"})
            continue

        target_dir = Path(template["sourceDir"])
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = _unique_target_path(target_dir, filename)
        with open(target_path, "wb") as f:
            f.write(content)
        saved.append(target_path.name)
        assignments.append({
            "filename": target_path.name,
            "templateId": resolved_template_id,
            "templateName": template.get("name") or template.get("label") or resolved_template_id,
        })

    if not saved and skipped:
        detail = "；".join(f"{item['filename']}：{item['reason']}" for item in skipped[:5])
        raise HTTPException(400, detail=f"没有可收录的模板文件。{detail}")

    templates = sync_template_library(enrich_structures=True)
    return {
        "success": True,
        "templateId": template_id or "",
        "saved": saved,
        "assignments": assignments,
        "skipped": skipped,
        "count": len(saved),
        "templates": [public_template_view(item) for item in templates],
    }


@app.get("/api/ai/model-config")
async def get_ai_model_config():
    return {"success": True, **get_model_config(mask_secret=True)}


@app.post("/api/ai/model-config")
async def update_ai_model_config(payload: AiModelConfigRequest):
    config = save_model_config(
        request_url=payload.request_url,
        api_key=payload.api_key,
        model_name=payload.model_name,
        models=payload.models,
    )
    return {"success": True, **config}


@app.post("/api/ai/generate")
async def generate_ai_document(
    files: list[UploadFile] = File(default=[]),
    template_files: list[UploadFile] = File(default=[]),
    material_files: list[UploadFile] = File(default=[]),
    prompt: str = Form(default=""),
    request_url: str = Form(default=""),
    api_key: str = Form(default=""),
    model_name: str = Form(default=""),
    template_id: str | None = Form(default=None),
    temperature: float = Form(default=0.2),
    speed_mode: str = Form(default="standard"),
    strict_reference_isolation: str = Form(default="false"),
):
    """根据上传材料、用户要求和模板库生成带公文格式的 docx。"""
    model_config = get_model_config()
    resolved_request_url = request_url.strip() or str(model_config.get("requestUrl") or "")
    resolved_api_key = api_key.strip()
    if not resolved_api_key or set(resolved_api_key) <= {"*"}:
        resolved_api_key = str(model_config.get("apiKey") or "")
    resolved_model_name = model_name.strip() or str(model_config.get("modelName") or "")
    if not resolved_api_key:
        raise HTTPException(400, detail="模型 API Key 未配置，请先在模型设置中填写并保存")

    saved_paths: list[Path] = []
    saved_template_paths: list[Path] = []
    # files 是旧客户端字段，继续作为业务材料兼容。
    effective_material_files = material_files or files
    if len(effective_material_files) > MAX_UPLOAD_BATCH:
        raise HTTPException(400, detail=f"单次最多上传 {MAX_UPLOAD_BATCH} 个材料文件")
    strict_mode = str(strict_reference_isolation).strip().lower() in {"1", "true", "yes", "on"}
    try:
        for file in [*template_files, *effective_material_files]:
            filename = Path(file.filename or "").name
            suffix = Path(filename).suffix.lower()
            if suffix in {".wps", ".ofd"}:
                raise HTTPException(
                    400,
                    detail=f"{filename or '文件'} 暂不支持直接解析，请先转换为 DOCX 或 PDF 后再上传",
                )
            if suffix not in SUPPORTED_INPUT_SUFFIXES:
                raise HTTPException(
                    400,
                    detail=f"{filename or '文件'} 格式不支持，请上传 DOCX、TXT、PDF 或图片文件",
                )
            content = await file.read()
            if len(content) > MAX_AI_UPLOAD_SIZE:
                raise HTTPException(400, detail=f"{filename} 超过 30MB")
            fpath = build_upload_temp_path(TEMP_DIR, filename, suffix)
            with open(fpath, "wb") as f:
                f.write(content)
            (saved_template_paths if file in template_files else saved_paths).append(fpath)

        def export_to_converter(documents: list[dict]) -> list[dict]:
            for document in documents:
                synced_file = _register_ai_document_for_converter(document)
                document["fileId"] = synced_file["id"]
                generated_ai_documents[document["documentId"]] = document
            _persist_ai_document_registry()
            return documents

        result = run_workflow(
            prompt=prompt,
            material_paths=saved_paths,
            template_paths=saved_template_paths,
            request_url=resolved_request_url,
            api_key=resolved_api_key,
            model_name=resolved_model_name,
            template_id=template_id,
            temperature=temperature,
            speed_mode=speed_mode,
            strict_reference_isolation=strict_mode,
            export_documents=export_to_converter,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, detail=f"AI 公文生成失败：{str(e)}")

    if not result.get("success"):
        raise HTTPException(500, detail=result.get("error") or "公文生成流程失败")
    synced_files = [{
        "id": document["fileId"], "name": document.get("filename"),
        "size": Path(document["path"]).stat().st_size, "templateId": document.get("templateId"),
        "documentId": document.get("documentId"),
    } for document in result.get("documents", [result])]
    attach_export_records(result.get("taskId", ""), result.get("documents", []))
    _persist_ai_document_registry()
    document_id = result["documentId"]
    first_document = (result.get("documents") or [result])[0] if result.get("documents") is not None else result
    return {
        "success": True,
        "documentId": document_id,
        "fileId": synced_files[0]["id"] if synced_files else "",
        "filename": result["filename"],
        "templateId": result["templateId"],
        "templateName": result["templateName"],
        "reference": result.get("reference"),
        "preview": result["preview"],
        "paragraphs": result["paragraphs"],
        "documents": result.get("documents", [result]),
        "files": synced_files,
        "warnings": result["warnings"],
        "readReport": result.get("readReport"),
        "processingMode": result.get("processingMode") or first_document.get("processingMode"),
        "draftLength": result.get("draftLength", first_document.get("draftLength", 0)),
        "summaryLength": result.get("summaryLength", first_document.get("summaryLength", 0)),
        "fullTextLength": result.get("fullTextLength", first_document.get("fullTextLength", 0)),
        "stages": result.get("stages") or first_document.get("stages"),
        "generation": result.get("generation") or first_document.get("generation"),
        "timings": result.get("timings"),
        "batchMode": result.get("batchMode"),
        "batchHint": result.get("batchHint"),
        "taskId": result.get("taskId"),
        "workflowStatus": result.get("status"),
        "nodes": result.get("nodes"),
        "failures": result.get("failures") or [],
    }


@app.get("/api/ai/tasks/{task_id}")
async def get_ai_task(task_id: str):
    task = get_workflow_task(task_id)
    if not task:
        raise HTTPException(404, detail="任务不存在或已过期")
    return {"success": True, **task}


@app.get("/api/ai/download/{document_id}")
async def download_ai_document(document_id: str):
    info = generated_ai_documents.get(document_id)
    if not info:
        _load_ai_document_registry()
        info = generated_ai_documents.get(document_id)
    if not info:
        raise HTTPException(404, detail="AI 生成结果不存在或已过期，请重新生成")
    # 优先返回工作台最新结果，避免局部编辑后双通道不一致
    file_id = str(info.get("fileId") or "")
    file_info = uploaded_files.get(file_id) if file_id else None
    path = Path(file_info["result_path"]) if file_info and file_info.get("result_path") else Path(info["path"])
    filename = (file_info or {}).get("name") or info.get("filename") or "AI公文.docx"
    if not path.exists():
        raise HTTPException(404, detail="AI 生成文件不存在或已过期，请重新生成")
    return FileResponse(
        path=path,
        filename=filename if str(filename).lower().endswith(".docx") else f"{Path(filename).stem}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/api/upload")
async def upload(files: list[UploadFile] = File(...)):
    if len(files) > MAX_UPLOAD_BATCH:
        raise HTTPException(400, detail=f"单次最多上传 {MAX_UPLOAD_BATCH} 个文件")
    result = []
    skipped = []
    for file in files:
        filename = Path(file.filename or "").name
        if not filename or not filename.lower().endswith(".docx"):
            skipped.append({
                "filename": filename or "(unknown)",
                "reason": "仅支持 .docx 文件",
            })
            continue
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            skipped.append({
                "filename": filename,
                "reason": f"超过 {MAX_FILE_SIZE // (1024 * 1024)}MB 上限",
            })
            continue
        if not content:
            skipped.append({"filename": filename, "reason": "文件为空"})
            continue
        fid = uuid.uuid4().hex
        fpath = TEMP_DIR / f"{fid}.docx"
        with open(fpath, "wb") as f:
            f.write(content)
        uploaded_files[fid] = {
            "path": fpath, "name": filename, "size": len(content),
            "processed": False, "result_id": None, "roles": {},
            "template_id": get_default_template_id(),
            "log": [], "undo_stack": [],
        }
        result.append({"id": fid, "name": filename, "size": len(content)})
    if not result and skipped:
        detail = "；".join(f"{item['filename']}：{item['reason']}" for item in skipped[:5])
        raise HTTPException(400, detail=f"没有可上传的文件。{detail}")
    return {"success": True, "files": result, "skipped": skipped}


@app.get("/api/structure/{file_id}")
async def get_structure(file_id: str, template_id: str | None = None):
    """识别结构清单（带置信度），供人工校正。"""
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    resolved_template_id = _validate_template_id(template_id or info.get("template_id"))
    try:
        items = build_structure(str(info["path"]), resolved_template_id)
    except Exception as e:
        raise HTTPException(500, detail=f"结构识别失败：{str(e)}")
    info["template_id"] = resolved_template_id
    return {
        "success": True, "name": info["name"],
        "templateId": resolved_template_id, "paragraphs": items,
    }


@app.post("/api/convert/{file_id}")
async def convert_one(file_id: str, payload: ConvertRequest | None = None):
    """按人工校正后的类型套格式。payload.roles 缺失则用自动识别结果。"""
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")

    roles = (payload.roles if payload else None) or {}
    template_id = _validate_template_id(
        (payload.template_id if payload else None) or info.get("template_id")
    )
    result_id = uuid.uuid4().hex
    output_path = TEMP_DIR / f"{result_id}.docx"

    try:
        result = convert_with_roles(
            str(info["path"]), str(output_path),
            roles=roles, template_id=template_id,
        )
    except Exception as e:
        raise HTTPException(500, detail=f"转换失败：{str(e)}")

    _clear_result_artifacts(info)
    previous_result_id = info.get("result_id")
    info.update({
        "processed": True, "result_id": result_id,
        "result_path": str(output_path), "roles": roles, "template_id": template_id,
        "log": result["log"], "undo_stack": [],
    })
    # 同步 AI 下载通道，避免助手下载仍指向旧文件
    linked_document = None
    linked_document_id = None
    for document_id, document in list(generated_ai_documents.items()):
        if str(document.get("fileId") or "") == file_id or document_id == previous_result_id:
            linked_document = document
            linked_document_id = document_id
            break
    if linked_document is not None:
        linked_document["path"] = str(output_path)
        linked_document["documentId"] = result_id
        linked_document["filename"] = info.get("name") or linked_document.get("filename")
        linked_document["fileId"] = file_id
        if linked_document_id != result_id:
            generated_ai_documents.pop(linked_document_id, None)
        generated_ai_documents[result_id] = linked_document
    _persist_ai_document_registry()
    stem = Path(info["name"]).stem
    download_name = (
        f"{stem}.docx"
        if stem.endswith("_AI公文") or stem.endswith("_公文格式")
        else f"{stem}_公文格式.docx"
    )
    return {
        "success": True, "fileId": file_id, "resultId": result_id,
        "templateId": template_id,
        "filename": download_name,
        "log": result["log"], "message": "转换完成",
    }


@app.get("/api/preview-result/{file_id}")
async def preview_result(file_id: str):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    rpath = info["result_path"]
    if not Path(rpath).exists():
        raise HTTPException(404, detail="处理结果文件不存在")
    items = build_result_preview(rpath, info.get("roles"), info.get("template_id"))
    return {"success": True, "name": info["name"], "type": "processed", "paragraphs": items}


@app.get("/api/render-check/{file_id}")
async def render_check(file_id: str):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    rpath = Path(info["result_path"])
    if not rpath.exists():
        raise HTTPException(404, detail="处理结果文件不存在")

    result = render_docx_to_pdf(rpath)
    if result.get("status") == "ok":
        info["render_pdf_path"] = result["pdfPath"]
    return {
        "success": result.get("status") == "ok",
        "status": result.get("status"),
        "message": result.get("message", "真实渲染完成"),
        "cached": result.get("cached", False),
        "pdfAvailable": result.get("status") == "ok",
        "pdfDownloadUrl": f"/api/render-pdf/{file_id}" if result.get("status") == "ok" else None,
        "detail": result.get("detail"),
    }


@app.get("/api/render-pdf/{file_id}")
async def render_pdf(file_id: str):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("render_pdf_path"):
        raise HTTPException(404, detail="渲染 PDF 不存在，请先执行真实渲染校验")
    pdf_path = Path(info["render_pdf_path"])
    if not pdf_path.exists():
        raise HTTPException(404, detail="渲染 PDF 不存在，请先执行真实渲染校验")
    return FileResponse(
        path=pdf_path,
        filename=f"{Path(info['name']).stem}_真实渲染预览.pdf",
        media_type="application/pdf",
    )


@app.post("/api/edit/{file_id}")
async def edit_result(file_id: str, payload: EditRequest):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")

    rpath = Path(info["result_path"])
    if not rpath.exists():
        original = info.get("path")
        if original and original.exists():
            convert_with_roles(
                str(original), str(rpath), roles=info.get("roles") or {},
                template_id=info.get("template_id"),
            )
        else:
            raise HTTPException(404, detail="处理结果文件不存在，请重新上传并处理")

    undo_path = TEMP_DIR / f"{uuid.uuid4().hex}_undo.docx"
    shutil.copyfile(rpath, undo_path)
    undo_stack = info.setdefault("undo_stack", [])
    undo_stack.append(str(undo_path))
    if len(undo_stack) > 10:
        old = Path(undo_stack.pop(0))
        if old.exists():
            old.unlink()

    try:
        apply_local_edit(
            str(rpath), payload.selection, payload.font, payload.paragraph,
            template_id=info.get("template_id"),
        )
    except Exception as e:
        if undo_path.exists():
            shutil.copyfile(undo_path, rpath)
            undo_path.unlink()
        if undo_stack and undo_stack[-1] == str(undo_path):
            undo_stack.pop()
        raise HTTPException(500, detail=f"应用修改失败：{str(e)}")

    _delete_file(info.get("render_pdf_path"))
    info["render_pdf_path"] = None
    _persist_ai_document_registry()
    return {
        "success": True,
        "paragraphs": build_result_preview(
            str(rpath), info.get("roles"), info.get("template_id"),
        ),
    }


@app.post("/api/undo/{file_id}")
async def undo_edit(file_id: str):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    undo_stack = info.get("undo_stack") or []
    if not undo_stack:
        raise HTTPException(400, detail="没有可撤销的修改")
    rpath = Path(info["result_path"])
    undo_path = Path(undo_stack.pop())
    if not undo_path.exists():
        raise HTTPException(404, detail="撤销文件不存在")
    shutil.copyfile(undo_path, rpath)
    undo_path.unlink()
    _delete_file(info.get("render_pdf_path"))
    info["render_pdf_path"] = None
    _persist_ai_document_registry()
    return {
        "success": True,
        "paragraphs": build_result_preview(
            str(rpath), info.get("roles"), info.get("template_id"),
        ),
    }


@app.get("/api/download/{file_id}")
async def download(file_id: str):
    info = _get_uploaded_file(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期，请重新上传或重新生成")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    rpath = info["result_path"]
    if not Path(rpath).exists():
        raise HTTPException(404, detail="处理结果文件不存在或已过期，请重新处理")
    name = str(info["name"] or "公文.docx")
    stem = Path(name).stem
    # AI 成品已带业务后缀时不再叠加“_公文格式”
    if stem.endswith("_AI公文") or stem.endswith("_公文格式"):
        download_name = f"{stem}.docx"
    else:
        download_name = f"{stem}_公文格式.docx"
    return FileResponse(
        path=rpath,
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.delete("/api/files/{file_id}")
async def delete_one_file(file_id: str):
    if not _purge_file_record(file_id):
        raise HTTPException(404, detail="文件不存在或已过期")
    return {"success": True, "fileId": file_id}


@app.delete("/api/files")
async def clear_files():
    for file_id in list(uploaded_files.keys()):
        _purge_file_record(file_id)
    generated_ai_documents.clear()
    _persist_ai_document_registry()
    return {"success": True}


if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
