"""FastAPI 后端入口 —— Word 公文格式转换服务

流程：上传 → 识别结构(/structure) → 人工校正 → 按校正套格式(/convert) → 预览/下载
"""

import uuid
import shutil
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
    generate_document,
    generate_documents,
    get_template_by_id,
    get_model_config,
    save_model_config,
    scan_template_library,
    sync_template_library,
)

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
STATIC_DIR = BASE_DIR / "static"
MAX_FILE_SIZE = 20 * 1024 * 1024

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
        if f.is_file() and (now - f.stat().st_mtime) > max_age_minutes * 60:
            try:
                f.unlink()
            except OSError:
                pass
    expired = [k for k, v in list(uploaded_files.items())
               if v["path"] and not v["path"].exists()]
    for k in expired:
        uploaded_files.pop(k, None)


@asynccontextmanager
async def lifespan(app: FastAPI):
    cleanup_old_files()
    sync_template_library()
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
    _delete_file(info.get("result_path"))
    for undo_path in info.get("undo_stack") or []:
        _delete_file(undo_path)
    info["undo_stack"] = []


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
    return {"success": True, "templates": scan_template_library()}


@app.post("/api/ai/templates/sync")
async def sync_ai_templates():
    templates = sync_template_library()
    return {"success": True, "count": len(templates), "templates": templates}


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
    if not template_id:
        raise HTTPException(400, detail="请先选择一个具体公文模板类别")
    template = get_template_by_id(template_id)
    if not template:
        raise HTTPException(400, detail="未找到对应模板类别")

    target_dir = Path(template["sourceDir"])
    target_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for file in files:
        filename = Path(file.filename or "").name
        suffix = Path(filename).suffix.lower()
        if not filename or filename.startswith("~$") or suffix not in REFERENCE_SUFFIXES:
            continue
        content = await file.read()
        if len(content) > MAX_AI_UPLOAD_SIZE:
            raise HTTPException(400, detail=f"{filename} 超过 30MB")
        target_path = _unique_target_path(target_dir, filename)
        with open(target_path, "wb") as f:
            f.write(content)
        saved.append(target_path.name)

    templates = sync_template_library()
    return {
        "success": True,
        "templateId": template_id,
        "saved": saved,
        "count": len(saved),
        "templates": templates,
    }


@app.get("/api/ai/model-config")
async def get_ai_model_config():
    return {"success": True, **get_model_config()}


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
    prompt: str = Form(default=""),
    request_url: str = Form(default=""),
    api_key: str = Form(default=""),
    model_name: str = Form(default=""),
    template_id: str | None = Form(default=None),
    temperature: float = Form(default=0.2),
):
    """根据上传材料、用户要求和模板库生成带公文格式的 docx。"""
    saved_paths: list[Path] = []
    try:
        for file in files:
            suffix = Path(file.filename or "").suffix.lower()
            if suffix not in SUPPORTED_INPUT_SUFFIXES:
                continue
            content = await file.read()
            if len(content) > MAX_AI_UPLOAD_SIZE:
                raise HTTPException(400, detail=f"{file.filename} 超过 30MB")
            input_id = uuid.uuid4().hex
            fpath = TEMP_DIR / f"{input_id}{suffix}"
            with open(fpath, "wb") as f:
                f.write(content)
            saved_paths.append(fpath)

        result = generate_documents(
            prompt=prompt,
            upload_paths=saved_paths,
            request_url=request_url,
            api_key=api_key,
            model_name=model_name,
            template_id=template_id,
            temperature=temperature,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, detail=f"AI 公文生成失败：{str(e)}")

    synced_files = []
    for document in result.get("documents", [result]):
        generated_ai_documents[document["documentId"]] = document
        synced_file = _register_ai_document_for_converter(document)
        document["fileId"] = synced_file["id"]
        synced_files.append(synced_file)
    document_id = result["documentId"]
    return {
        "success": True,
        "documentId": document_id,
        "fileId": synced_files[0]["id"] if synced_files else "",
        "filename": result["filename"],
        "templateId": result["templateId"],
        "templateName": result["templateName"],
        "preview": result["preview"],
        "paragraphs": result["paragraphs"],
        "documents": result.get("documents", [result]),
        "files": synced_files,
        "warnings": result["warnings"],
    }


@app.get("/api/ai/download/{document_id}")
async def download_ai_document(document_id: str):
    info = generated_ai_documents.get(document_id)
    if not info:
        raise HTTPException(404, detail="AI 生成结果不存在或已过期")
    path = Path(info["path"])
    if not path.exists():
        raise HTTPException(404, detail="AI 生成文件不存在")
    return FileResponse(
        path=path,
        filename=info["filename"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/api/upload")
async def upload(files: list[UploadFile] = File(...)):
    result = []
    for file in files:
        if not file.filename or not file.filename.lower().endswith(".docx"):
            continue
        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            continue
        fid = uuid.uuid4().hex
        fpath = TEMP_DIR / f"{fid}.docx"
        with open(fpath, "wb") as f:
            f.write(content)
        uploaded_files[fid] = {
            "path": fpath, "name": file.filename, "size": len(content),
            "processed": False, "result_id": None, "roles": {},
            "template_id": get_default_template_id(),
            "log": [], "undo_stack": [],
        }
        result.append({"id": fid, "name": file.filename, "size": len(content)})
    return {"success": True, "files": result}


@app.get("/api/structure/{file_id}")
async def get_structure(file_id: str, template_id: str | None = None):
    """识别结构清单（带置信度），供人工校正。"""
    info = uploaded_files.get(file_id)
    if not info or not info["path"].exists():
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
    info = uploaded_files.get(file_id)
    if not info or not info["path"].exists():
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
    info.update({
        "processed": True, "result_id": result_id,
        "result_path": str(output_path), "roles": roles, "template_id": template_id,
        "log": result["log"], "undo_stack": [],
    })
    return {
        "success": True, "fileId": file_id, "resultId": result_id,
        "templateId": template_id,
        "filename": f"{Path(info['name']).stem}_公文格式.docx",
        "log": result["log"], "message": "转换完成",
    }


@app.get("/api/preview-result/{file_id}")
async def preview_result(file_id: str):
    info = uploaded_files.get(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    rpath = info["result_path"]
    if not Path(rpath).exists():
        raise HTTPException(404, detail="处理结果文件不存在")
    items = build_result_preview(rpath, info.get("roles"), info.get("template_id"))
    return {"success": True, "name": info["name"], "type": "processed", "paragraphs": items}


@app.post("/api/edit/{file_id}")
async def edit_result(file_id: str, payload: EditRequest):
    info = uploaded_files.get(file_id)
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

    return {
        "success": True,
        "paragraphs": build_result_preview(
            str(rpath), info.get("roles"), info.get("template_id"),
        ),
    }


@app.post("/api/undo/{file_id}")
async def undo_edit(file_id: str):
    info = uploaded_files.get(file_id)
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
    return {
        "success": True,
        "paragraphs": build_result_preview(
            str(rpath), info.get("roles"), info.get("template_id"),
        ),
    }


@app.get("/api/download/{file_id}")
async def download(file_id: str):
    info = uploaded_files.get(file_id)
    if not info:
        raise HTTPException(404, detail="文件不存在或已过期")
    if not info.get("processed") or not info.get("result_path"):
        raise HTTPException(400, detail="该文件尚未处理")
    rpath = info["result_path"]
    if not Path(rpath).exists():
        raise HTTPException(404, detail="处理结果文件不存在，请重新处理")
    stem = Path(info["name"]).stem
    return FileResponse(
        path=rpath,
        filename=f"{stem}_公文格式.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.delete("/api/files")
async def clear_files():
    for info in uploaded_files.values():
        try:
            if info["path"] and info["path"].exists():
                info["path"].unlink()
        except OSError:
            pass
        _clear_result_artifacts(info)
    uploaded_files.clear()
    return {"success": True}


if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
