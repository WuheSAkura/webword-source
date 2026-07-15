"""FastAPI 后端入口 —— Word 公文格式转换服务

流程：上传 → 识别结构(/structure) → 人工校正 → 按校正套格式(/convert) → 预览/下载
"""

import uuid
import shutil
from contextlib import asynccontextmanager
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from docx import Document
from docx.oxml.ns import qn

from config import get_config, ROLE_LABELS
from converter import (
    convert_with_roles,
    build_structure,
    classify_one,
    split_paragraph_text,
    apply_local_edit,
)

BASE_DIR = Path(__file__).resolve().parent
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)
STATIC_DIR = BASE_DIR / "static"
MAX_FILE_SIZE = 20 * 1024 * 1024

# { file_id: {path, name, size, processed, result_id, result_path, roles, log, undo_stack} }
uploaded_files = {}
processed_files = {}


class ConvertRequest(BaseModel):
    roles: dict[str, str] | None = None  # {逻辑段索引(str): 角色}


class EditRequest(BaseModel):
    selection: dict
    font: dict | None = None
    paragraph: dict | None = None


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
    yield


app = FastAPI(title="Word公文格式转换工具", version="2.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def build_result_preview(file_path: str) -> list[dict]:
    """从套好格式的 docx 提取段落预览（含 run 字体），供结果预览与局部编辑。"""
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

    items = []
    for i, (source_index, paragraph, text) in enumerate(logical_items):
        role, _ = classify_one(text, i, total)
        align_map = {None: "左", 0: "左", 1: "中", 2: "右", 3: "两端"}
        items.append({
            "index": i,
            "text": text[:200] + ("..." if len(text) > 200 else ""),
            "fullText": text,
            "role": role,
            "sourceIndex": source_index,
            "roleLabel": ROLE_LABELS.get(role, "正文"),
            "runs": extract_runs(paragraph, text),
            "align": align_map.get(paragraph.alignment, "左"),
        })
    return items


# ── API ──
@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.get("/api/config")
async def get_format_config():
    """返回格式配置与角色标签，供前端校正/编辑面板使用（消除前端硬编码）。"""
    config = get_config()
    return {"success": True, "styles": config["styles"], "labels": ROLE_LABELS}


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
            "log": [], "undo_stack": [],
        }
        result.append({"id": fid, "name": file.filename, "size": len(content)})
    return {"success": True, "files": result}


@app.get("/api/structure/{file_id}")
async def get_structure(file_id: str):
    """识别结构清单（带置信度），供人工校正。"""
    info = uploaded_files.get(file_id)
    if not info or not info["path"].exists():
        raise HTTPException(404, detail="文件不存在或已过期")
    try:
        items = build_structure(str(info["path"]))
    except Exception as e:
        raise HTTPException(500, detail=f"结构识别失败：{str(e)}")
    return {"success": True, "name": info["name"], "paragraphs": items}


@app.post("/api/convert/{file_id}")
async def convert_one(file_id: str, payload: ConvertRequest | None = None):
    """按人工校正后的类型套格式。payload.roles 缺失则用自动识别结果。"""
    info = uploaded_files.get(file_id)
    if not info or not info["path"].exists():
        raise HTTPException(404, detail="文件不存在或已过期")

    roles = (payload.roles if payload else None) or {}
    result_id = uuid.uuid4().hex
    output_path = TEMP_DIR / f"{result_id}.docx"

    try:
        result = convert_with_roles(str(info["path"]), str(output_path), roles=roles)
    except Exception as e:
        raise HTTPException(500, detail=f"转换失败：{str(e)}")

    info.update({
        "processed": True, "result_id": result_id,
        "result_path": str(output_path), "roles": roles,
        "log": result["log"], "undo_stack": [],
    })
    processed_files[result_id] = str(output_path)
    return {
        "success": True, "fileId": file_id, "resultId": result_id,
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
    items = build_result_preview(rpath)
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
            convert_with_roles(str(original), str(rpath), roles=info.get("roles") or {})
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
        apply_local_edit(str(rpath), payload.selection, payload.font, payload.paragraph)
    except Exception as e:
        if undo_path.exists():
            shutil.copyfile(undo_path, rpath)
        raise HTTPException(500, detail=f"应用修改失败：{str(e)}")

    return {"success": True, "paragraphs": build_result_preview(str(rpath))}


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
    return {"success": True, "paragraphs": build_result_preview(str(rpath))}


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
        if info.get("result_path"):
            try:
                Path(info["result_path"]).unlink()
            except OSError:
                pass
    uploaded_files.clear()
    processed_files.clear()
    return {"success": True}


if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
