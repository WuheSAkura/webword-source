"""AI document writing service helpers.

The implementation is intentionally OpenAI-compatible so it can point at an
intranet model gateway. Users provide request_url, api_key and model_name from
the assistant panel instead of relying on hard-coded cloud settings.
"""

from __future__ import annotations

import json
import os
import sqlite3
import uuid
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from config import get_template_catalog, normalize_template_id
from converter import build_structure, convert_with_roles

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
TEMPLATE_SOURCE_DIR = (
    PROJECT_DIR
    / "20260623--整理汇总常用公文及范例（环食药侦）"
    / "2-1附件1：15种常用公文（原件）"
)
TEMP_DIR = BASE_DIR / "temp"
ENV_PATH = BASE_DIR / ".env"
TEMPLATE_DB_PATH = BASE_DIR / "template_library.db"
MAX_AI_UPLOAD_SIZE = 30 * 1024 * 1024
DIRECT_DRAFT_CHAR_LIMIT = 12000
SUMMARY_TRIGGER_CHAR_LIMIT = 12000
SUMMARY_CHUNK_CHAR_LIMIT = 7000
SUMMARY_CHUNK_OVERLAP = 350
MAX_AGGREGATED_SUMMARY_CHARS = 22000
READ_REPORT_SAMPLE_LIMIT = 20

SUPPORTED_INPUT_SUFFIXES = {".docx", ".txt"}
REFERENCE_SUFFIXES = {".docx", ".txt", ".wps"}
SUMMARY_FIRST_TEMPLATE_IDS = {"minutes"}

LEGACY_DOCUMENT_TYPE_HINTS = [
    ("request", ("请示",)),
    ("report", ("报告", "汇报")),
    ("notice", ("通知",)),
    ("bulletin", ("通报",)),
    ("letter", ("函", "商请", "邀请")),
    ("reply", ("批复",)),
    ("opinion", ("意见",)),
    ("proposal", ("议案",)),
    ("minutes", ("纪要", "会议纪要")),
    ("announcement", ("公告",)),
    ("public_notice", ("通告",)),
    ("decision", ("决定",)),
    ("resolution", ("决议",)),
    ("order", ("命令", "令")),
    ("communique", ("公报",)),
]


DOCUMENT_TYPE_HINTS = [
    ("request", ("\u8bf7\u793a",)),
    ("report", ("\u62a5\u544a", "\u6c47\u62a5")),
    ("notice", ("\u901a\u77e5",)),
    ("bulletin", ("\u901a\u62a5",)),
    ("letter", ("\u51fd", "\u5546\u8bf7", "\u9080\u8bf7")),
    ("reply", ("\u6279\u590d",)),
    ("opinion", ("\u610f\u89c1",)),
    ("proposal", ("\u8bae\u6848",)),
    ("minutes", ("\u7eaa\u8981", "\u4f1a\u8bae\u7eaa\u8981", "\u4f1a\u8bae\u603b\u7ed3", "\u603b\u7ed3\u7eaa\u8981")),
    ("announcement", ("\u516c\u544a",)),
    ("public_notice", ("\u901a\u544a",)),
    ("decision", ("\u51b3\u5b9a",)),
    ("resolution", ("\u51b3\u8bae",)),
    ("order", ("\u547d\u4ee4", "\u4ee4")),
    ("communique", ("\u516c\u62a5",)),
]


def _template_sort_key(item: dict[str, Any]) -> tuple[int, str]:
    order_by_id = {entry["id"]: entry["order"] for entry in get_template_catalog()}
    return (int(order_by_id.get(item.get("id"), 999)), str(item.get("name") or ""))


def load_local_env() -> None:
    if not ENV_PATH.exists():
        return
    for line in ENV_PATH.read_text(encoding="utf-8", errors="ignore").splitlines():
        text = line.strip()
        if not text or text.startswith("#") or "=" not in text:
            continue
        key, value = text.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def get_model_config() -> dict[str, Any]:
    load_local_env()
    models = [
        item.strip()
        for item in os.getenv("AI_AVAILABLE_MODELS", "deepseek-chat,deepseek-reasoner").split(",")
        if item.strip()
    ]
    current_model = os.getenv("DEEPSEEK_MODEL", models[0] if models else "deepseek-chat")
    if current_model not in models:
        models.insert(0, current_model)
    return {
        "requestUrl": os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
        "apiKey": os.getenv("DEEPSEEK_API_KEY", ""),
        "modelName": current_model,
        "models": models,
    }


def save_model_config(request_url: str, api_key: str, model_name: str, models: list[str] | None = None) -> dict[str, Any]:
    clean_models = [item.strip() for item in (models or []) if item and item.strip()]
    if model_name.strip() and model_name.strip() not in clean_models:
        clean_models.insert(0, model_name.strip())
    if not clean_models:
        clean_models = ["deepseek-chat", "deepseek-reasoner"]

    values = {
        "DEEPSEEK_BASE_URL": request_url.strip() or "https://api.deepseek.com",
        "DEEPSEEK_API_KEY": api_key.strip(),
        "DEEPSEEK_MODEL": model_name.strip() or clean_models[0],
        "AI_AVAILABLE_MODELS": ",".join(clean_models),
    }
    ENV_PATH.write_text("\n".join(f"{key}={value}" for key, value in values.items()) + "\n", encoding="utf-8")
    for key, value in values.items():
        os.environ[key] = value
    return get_model_config()


def legacy_extract_text(path: Path, limit: int = 12000) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        doc = Document(str(path))
        text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    elif suffix == ".txt":
        text = path.read_text(encoding="utf-8", errors="ignore")
    else:
        text = f"[暂不支持直接抽取 {suffix} 文件正文，请先转为 docx 或 txt]"
    text = text.strip()
    return text[:limit] + ("\n...[内容已截断]" if len(text) > limit else "")


def _new_stats() -> dict[str, int]:
    return {
        "bodyParagraphs": 0,
        "tables": 0,
        "tableRows": 0,
        "tableCells": 0,
        "headers": 0,
        "footers": 0,
        "images": 0,
        "chars": 0,
        "items": 0,
    }


def _add_inventory_item(
    items: list[dict[str, Any]],
    stats: dict[str, int],
    source_type: str,
    location: str,
    text: str,
) -> None:
    clean = " ".join(text.replace("\r", "\n").split()) if source_type == "table" else text.strip()
    if not clean:
        return
    index = len(items) + 1
    item = {
        "id": f"{source_type}.{index}",
        "sourceType": source_type,
        "location": location,
        "text": clean,
        "charCount": len(clean),
    }
    items.append(item)
    stats["chars"] += item["charCount"]
    stats["items"] = len(items)


def _iter_body_blocks(doc):
    table_index = 0
    paragraph_index = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph_index += 1
            yield "paragraph", paragraph_index, Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            table_index += 1
            yield "table", table_index, Table(child, doc)


def _read_header_footer(container, source_type: str, section_index: int, variant: str,
                        items: list[dict[str, Any]], stats: dict[str, int]) -> None:
    seen = set()
    for index, paragraph in enumerate(container.paragraphs, start=1):
        text = paragraph.text.strip()
        if not text or text in seen:
            continue
        seen.add(text)
        _add_inventory_item(
            items,
            stats,
            source_type,
            f"section {section_index} {variant} paragraph {index}",
            text,
        )
        stats["headers" if source_type == "header" else "footers"] += 1


def _collect_docx_inventory(path: Path) -> dict[str, Any]:
    doc = Document(str(path))
    items: list[dict[str, Any]] = []
    stats = _new_stats()
    warnings: list[str] = []

    for block_type, block_index, block in _iter_body_blocks(doc):
        if block_type == "paragraph":
            text = block.text.strip()
            if text:
                _add_inventory_item(items, stats, "body", f"paragraph {block_index}", text)
                stats["bodyParagraphs"] += 1
        elif block_type == "table":
            stats["tables"] += 1
            stats["tableRows"] += len(block.rows)
            for row_index, row in enumerate(block.rows, start=1):
                for cell_index, cell in enumerate(row.cells, start=1):
                    stats["tableCells"] += 1
                    cell_text = "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    if cell_text.strip():
                        _add_inventory_item(
                            items,
                            stats,
                            "table",
                            f"table {block_index} row {row_index} cell {cell_index}",
                            cell_text,
                        )

    for section_index, section in enumerate(doc.sections, start=1):
        for variant, header in (
            ("default", section.header),
            ("first", section.first_page_header),
            ("even", section.even_page_header),
        ):
            _read_header_footer(header, "header", section_index, variant, items, stats)
        for variant, footer in (
            ("default", section.footer),
            ("first", section.first_page_footer),
            ("even", section.even_page_footer),
        ):
            _read_header_footer(footer, "footer", section_index, variant, items, stats)

    stats["images"] = len(doc.inline_shapes)
    if stats["images"]:
        warnings.append(f"检测到 {stats['images']} 张图片，当前未抽取图片内文字；如材料是扫描件，生成结果可能遗漏图片内容。")
    if not items and stats["images"]:
        warnings.append("文档未抽取到可读文本且包含图片，疑似扫描件或图片版材料。")

    return {"items": items, "stats": stats, "warnings": warnings}


def _collect_txt_inventory(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="ignore")
    items: list[dict[str, Any]] = []
    stats = _new_stats()
    for index, line in enumerate(text.splitlines(), start=1):
        if line.strip():
            _add_inventory_item(items, stats, "text", f"line {index}", line)
    return {"items": items, "stats": stats, "warnings": []}


def extract_document_inventory(path: Path) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _collect_docx_inventory(path)
    if suffix == ".txt":
        return _collect_txt_inventory(path)

    stats = _new_stats()
    warning = f"Unsupported input file type: {suffix}. Convert it to docx or txt first."
    return {"items": [], "stats": stats, "warnings": [warning]}


def inventory_to_text(inventory: dict[str, Any]) -> str:
    lines = []
    for item in inventory.get("items", []):
        lines.append(f"[{item['id']} | {item['sourceType']} | {item['location']}]\n{item['text']}")
    return "\n\n".join(lines).strip()


def build_read_report(inventory: dict[str, Any], chunks: list[str] | None = None) -> dict[str, Any]:
    stats = dict(inventory.get("stats") or _new_stats())
    sample_items = []
    for item in inventory.get("items", [])[:READ_REPORT_SAMPLE_LIMIT]:
        sample_items.append({
            "id": item["id"],
            "sourceType": item["sourceType"],
            "location": item["location"],
            "charCount": item["charCount"],
            "text": item["text"][:120] + ("..." if len(item["text"]) > 120 else ""),
        })
    return {
        "stats": stats,
        "warnings": list(inventory.get("warnings") or []),
        "sampleItems": sample_items,
        "chunkCount": len(chunks or []),
        "chunkSources": [
            {
                "index": index,
                "charCount": len(chunk),
                "sourceIds": _source_ids_in_text(chunk),
            }
            for index, chunk in enumerate(chunks or [], start=1)
        ],
    }


def extract_text(path: Path, limit: int | None = DIRECT_DRAFT_CHAR_LIMIT) -> str:
    suffix = path.suffix.lower()
    if suffix in SUPPORTED_INPUT_SUFFIXES:
        text = inventory_to_text(extract_document_inventory(path))
    else:
        text = f"Unsupported input file type: {suffix}. Convert it to docx or txt first."
    text = text.strip()
    if limit is None or len(text) <= limit:
        return text
    return text[:limit] + "\n...[content truncated]"


def chunk_text(text: str, chunk_size: int = SUMMARY_CHUNK_CHAR_LIMIT,
               overlap: int = SUMMARY_CHUNK_OVERLAP) -> list[str]:
    clean = text.strip()
    if not clean:
        return []
    chunks = []
    start = 0
    while start < len(clean):
        end = min(len(clean), start + chunk_size)
        chunks.append(clean[start:end])
        if end >= len(clean):
            break
        start = max(end - overlap, start + 1)
    return chunks


def _source_ids_in_text(text: str) -> list[str]:
    ids = []
    for part in text.split("["):
        if "]" not in part:
            continue
        source_id = part.split("]", 1)[0].split("|", 1)[0].strip()
        if source_id and source_id not in ids:
            ids.append(source_id)
    return ids


def fallback_summary(text: str, limit: int = 1800) -> str:
    paragraphs = [p.strip() for p in text.splitlines() if p.strip()]
    if not paragraphs:
        return ""
    selected = paragraphs[:8]
    if len(paragraphs) > 12:
        selected.extend(paragraphs[-4:])
    summary = "\n".join(selected)
    return summary[:limit] + ("\n...[summary clipped]" if len(summary) > limit else "")


def summarize_chunk(
    text: str,
    prompt: str,
    template_label: str,
    request_url: str,
    api_key: str,
    model_name: str,
    index: int,
    total: int,
) -> str:
    system = (
        "You are a Chinese official-document analyst. Summarize the source "
        "faithfully for later official-document drafting. Preserve names, "
        "organizations, dates, locations, numbers, decisions, requirements, "
        "problems, responsibilities, and pending items. Do not invent facts."
    )
    user = (
        f"Target document type: {template_label or 'auto'}\n"
        f"User request: {prompt or 'Draft an official document from the material.'}\n"
        f"Source chunk {index}/{total}:\n{text}\n\n"
        "Return a dense Chinese summary with bullet-like short paragraphs. "
        "Keep all facts that may affect the final document. Preserve source "
        "ids such as [body.1], [table.4], [header.8] when mentioning facts."
    )
    return call_chat_model(
        request_url=request_url,
        api_key=api_key,
        model_name=model_name,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        temperature=0.1,
    )


def prepare_material(
    path: Path,
    prompt: str,
    template_label: str,
    request_url: str,
    api_key: str,
    model_name: str,
    force_summary: bool = False,
) -> tuple[dict[str, Any], list[str]]:
    name = display_name_from_temp_path(path)
    inventory = extract_document_inventory(path)
    full_text = inventory_to_text(inventory)
    warnings: list[str] = list(inventory.get("warnings") or [])
    chunks = chunk_text(full_text)
    should_summarize = force_summary or len(full_text) > SUMMARY_TRIGGER_CHAR_LIMIT or len(chunks) > 1
    read_report = build_read_report(inventory, chunks)

    if not should_summarize:
        return {
            "name": name,
            "text": full_text,
            "fullTextLength": len(full_text),
            "chunkCount": 1 if full_text else 0,
            "processingMode": "direct",
            "readReport": read_report,
        }, warnings

    chunk_summaries = []
    for index, chunk in enumerate(chunks, start=1):
        try:
            chunk_summaries.append(
                summarize_chunk(chunk, prompt, template_label, request_url, api_key, model_name, index, len(chunks))
            )
        except Exception as exc:
            warnings.append(f"{name} chunk {index} summary failed: {exc}")
            chunk_summaries.append(fallback_summary(chunk))

    summary_text = "\n\n".join(
        f"[Part {index}/{len(chunk_summaries)}]\n{summary}"
        for index, summary in enumerate(chunk_summaries, start=1)
        if summary.strip()
    )
    if len(summary_text) > MAX_AGGREGATED_SUMMARY_CHARS:
        try:
            summary_text = summarize_chunk(
                summary_text,
                prompt,
                template_label,
                request_url,
                api_key,
                model_name,
                1,
                1,
            )
        except Exception as exc:
            warnings.append(f"{name} aggregate summary failed: {exc}")
            summary_text = summary_text[:MAX_AGGREGATED_SUMMARY_CHARS] + "\n...[summary clipped]"

    return {
        "name": name,
        "text": summary_text,
        "fullTextLength": len(full_text),
        "chunkCount": len(chunks),
        "processingMode": "summary_first",
        "readReport": read_report,
    }, warnings


def scan_template_library() -> list[dict[str, Any]]:
    catalog_by_label = {item["label"]: item for item in get_template_catalog()}
    templates = []
    if not TEMPLATE_SOURCE_DIR.exists():
        return templates

    for folder in sorted([p for p in TEMPLATE_SOURCE_DIR.iterdir() if p.is_dir()], key=lambda p: p.name):
        files = [
            item for item in folder.iterdir()
            if item.is_file()
            and not item.name.startswith("~$")
            and item.suffix.lower() in REFERENCE_SUFFIXES
        ]
        sample_text = ""
        sample_file = ""
        for item in files:
            try:
                sample_text = extract_text(item, 1600)
                sample_file = item.name
                if sample_text and not sample_text.startswith("[暂不支持"):
                    break
            except Exception:
                continue
        template_id = guess_template_id(folder.name)
        catalog_item = next((v for v in catalog_by_label.values() if v["id"] == template_id), None)
        templates.append({
            "id": template_id,
            "name": folder.name,
            "label": catalog_item["label"] if catalog_item else folder.name,
            "fileCount": len(files),
            "sampleFile": sample_file,
            "sampleText": sample_text,
        })
    return sorted(templates, key=_template_sort_key)


def guess_template_id(text: str) -> str:
    for template_id, keywords in DOCUMENT_TYPE_HINTS:
        if any(keyword in text for keyword in keywords):
            return template_id
    return "generic"


def choose_template_id(prompt: str, file_texts: list[str], requested_template_id: str | None) -> str:
    if requested_template_id:
        return normalize_template_id(requested_template_id)
    combined = "\n".join([prompt, *file_texts])
    return normalize_template_id(guess_template_id(combined))


def init_template_db() -> None:
    TEMPLATE_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS document_templates (
                id TEXT NOT NULL,
                name TEXT NOT NULL,
                label TEXT NOT NULL,
                file_count INTEGER NOT NULL DEFAULT 0,
                source_dir TEXT NOT NULL,
                sample_file TEXT NOT NULL DEFAULT '',
                sample_path TEXT NOT NULL DEFAULT '',
                sample_text TEXT NOT NULL DEFAULT '',
                files_json TEXT NOT NULL DEFAULT '[]',
                updated_at TEXT NOT NULL,
                PRIMARY KEY (id, source_dir)
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_document_templates_id ON document_templates(id)")


def load_templates_from_db() -> list[dict[str, Any]]:
    init_template_db()
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, name, label, file_count, source_dir, sample_file,
                   sample_path, sample_text, files_json, updated_at
            FROM document_templates
            ORDER BY name
            """
        ).fetchall()
    templates = []
    for row in rows:
        try:
            files = json.loads(row["files_json"] or "[]")
        except json.JSONDecodeError:
            files = []
        templates.append({
            "id": row["id"],
            "name": row["name"],
            "label": row["label"],
            "fileCount": row["file_count"],
            "sourceDir": row["source_dir"],
            "sampleFile": row["sample_file"],
            "samplePath": row["sample_path"],
            "sampleText": row["sample_text"],
            "files": files,
            "updatedAt": row["updated_at"],
        })
    return sorted(templates, key=_template_sort_key)


def sync_template_library() -> list[dict[str, Any]]:
    init_template_db()
    catalog_by_label = {item["label"]: item for item in get_template_catalog()}
    templates = []
    if not TEMPLATE_SOURCE_DIR.exists():
        return load_templates_from_db()

    for folder in sorted([p for p in TEMPLATE_SOURCE_DIR.iterdir() if p.is_dir()], key=lambda p: p.name):
        files = [
            item for item in folder.iterdir()
            if item.is_file()
            and not item.name.startswith("~$")
            and item.suffix.lower() in REFERENCE_SUFFIXES
        ]
        sample_text = ""
        sample_file = ""
        for item in files:
            try:
                sample_text = extract_text(item, 1600)
                sample_file = item.name
                if sample_text and not sample_text.startswith("Unsupported input file type"):
                    break
            except Exception:
                continue
        template_id = guess_template_id(folder.name)
        catalog_item = next((v for v in catalog_by_label.values() if v["id"] == template_id), None)
        templates.append({
            "id": template_id,
            "name": folder.name,
            "label": catalog_item["label"] if catalog_item else folder.name,
            "fileCount": len(files),
            "sourceDir": str(folder),
            "sampleFile": sample_file,
            "samplePath": str(folder / sample_file) if sample_file else "",
            "sampleText": sample_text,
            "files": [item.name for item in files],
        })

    from datetime import datetime
    now = datetime.now().isoformat(timespec="seconds")
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.execute("DELETE FROM document_templates")
        conn.executemany(
            """
            INSERT INTO document_templates (
                id, name, label, file_count, source_dir, sample_file,
                sample_path, sample_text, files_json, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id, source_dir) DO UPDATE SET
                name=excluded.name,
                label=excluded.label,
                file_count=excluded.file_count,
                sample_file=excluded.sample_file,
                sample_path=excluded.sample_path,
                sample_text=excluded.sample_text,
                files_json=excluded.files_json,
                updated_at=excluded.updated_at
            """,
            [
                (
                    item["id"],
                    item["name"],
                    item["label"],
                    item["fileCount"],
                    item["sourceDir"],
                    item["sampleFile"],
                    item["samplePath"],
                    item["sampleText"],
                    json.dumps(item["files"], ensure_ascii=False),
                    now,
                )
                for item in templates
            ],
        )
    return sorted(load_templates_from_db(), key=_template_sort_key)


def scan_template_library() -> list[dict[str, Any]]:
    return sync_template_library()


def get_template_by_id(template_id: str | None) -> dict[str, Any] | None:
    if not template_id:
        return None
    return next((item for item in scan_template_library() if item["id"] == template_id), None)


def legacy_build_ai_messages(prompt: str, materials: list[dict[str, str]], template: dict | None) -> list[dict[str, str]]:
    template_note = ""
    if template:
        template_note = (
            f"参考文种模板：{template.get('name')}\n"
            f"参考文件：{template.get('sampleFile')}\n"
            f"模板摘录：\n{template.get('sampleText') or '无可抽取文本'}"
        )
    material_text = "\n\n".join(
        f"【材料 {index + 1}：{item['name']}】\n{item['text']}"
        for index, item in enumerate(materials)
    )
    system = (
        "你是一名熟悉党政机关和公安机关公文写作规范的中文写作助手。"
        "请严格依据用户材料和要求生成正式公文正文，不编造关键事实；"
        "缺失信息可使用规范占位符提示补充。输出纯文本，按公文段落换行，"
        "不要输出 Markdown 标题符号、代码块或解释性说明。"
    )
    user = (
        f"用户要求：\n{prompt or '请根据材料生成正式公文。'}\n\n"
        f"{template_note}\n\n"
        f"用户上传材料：\n{material_text or '无上传材料'}"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def build_ai_messages(prompt: str, materials: list[dict[str, Any]], template: dict | None) -> list[dict[str, str]]:
    template_note = ""
    if template:
        template_note = (
            f"Reference document type/template: {template.get('name')}\n"
            f"Reference file: {template.get('sampleFile')}\n"
            f"Reference excerpt:\n{template.get('sampleText') or 'No extractable reference text'}"
        )

    material_text = "\n\n".join(
        "\n".join([
            f"[Material {index + 1}: {item['name']}]",
            f"Processing mode: {item.get('processingMode', 'direct')}",
            f"Original length: {item.get('fullTextLength', len(item.get('text', '')))} chars",
            f"Chunks read: {item.get('chunkCount', 1)}",
            item.get("text", ""),
        ])
        for index, item in enumerate(materials)
    )
    template_id = template.get("id", "") if template else ""
    workflow_note = (
        "For meeting minutes or summary-minutes style documents, first convert the "
        "material into meeting facts: attendees/units, agenda, key discussion, "
        "decisions, tasks, owners, deadlines, and unresolved items; then draft the "
        "official document. For other document types, route the writing according "
        "to the selected or inferred document type. The final answer must be plain "
        "Chinese official-document text with paragraph line breaks only."
    )
    system = (
        "You are a Chinese official-document drafting assistant familiar with "
        "Party, government, and public-security document norms. Use only the "
        "provided material and user request; do not invent key facts. If required "
        "facts are missing, use formal placeholders for completion. Do not output "
        "Markdown, code fences, explanations, or analysis."
    )
    user = (
        f"User request:\n{prompt or 'Draft a formal official document from the material.'}\n\n"
        f"Resolved template id: {template_id or 'generic'}\n"
        f"{workflow_note}\n\n"
        f"{template_note}\n\n"
        f"Source material or full-document summaries:\n{material_text or 'No uploaded material.'}"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def resolve_chat_url(request_url: str) -> str:
    url = (request_url or "").strip().rstrip("/")
    if not url:
        raise ValueError("请填写模型请求地址")
    if url.endswith("/chat/completions"):
        return url
    if url.endswith("/v1"):
        return f"{url}/chat/completions"
    return f"{url}/v1/chat/completions"


def call_chat_model(
    request_url: str,
    api_key: str,
    model_name: str,
    messages: list[dict[str, str]],
    temperature: float = 0.2,
) -> str:
    load_local_env()
    request_url = request_url.strip() or os.getenv("DEEPSEEK_BASE_URL", "")
    api_key = api_key.strip() or os.getenv("DEEPSEEK_API_KEY", "")
    model_name = model_name.strip() or os.getenv("DEEPSEEK_MODEL", "")
    if not model_name:
        raise ValueError("请填写模型名称")
    payload = {
        "model": model_name,
        "messages": messages,
        "temperature": temperature,
        "stream": False,
    }
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(resolve_chat_url(request_url), data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"模型接口返回 {exc.code}: {detail[:500]}") from exc
    except Exception as exc:
        raise RuntimeError(f"模型接口请求失败: {exc}") from exc
    try:
        return body["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("模型接口响应格式不符合 OpenAI Chat Completions 规范") from exc


def fallback_draft(prompt: str, materials: list[dict[str, str]], template_label: str) -> str:
    title = prompt.strip().splitlines()[0][:48] if prompt.strip() else f"关于有关工作的{template_label}"
    snippets = []
    for item in materials:
        body = item["text"].replace("\n", " ")
        if body:
            snippets.append(f"{item['name']}：{body[:280]}")
    source = "\n".join(snippets) or "（请补充事实材料、时间地点、工作措施和办理要求。）"
    return (
        f"{title}\n\n"
        "有关单位：\n"
        f"根据工作需要，现就有关事项形成如下材料。\n\n"
        f"一、基本情况\n{source}\n\n"
        "二、工作意见\n请结合实际进一步核实完善相关情况，明确责任分工、时间节点和工作要求。\n\n"
        "三、有关要求\n请严格按照公文规范补充发文机关、主送机关、成文日期等要素后报审。"
    )


def write_plain_docx(text: str, output_path: Path) -> Path:
    doc = Document()
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        doc.add_paragraph(line.strip())
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def generate_document(
    prompt: str,
    upload_paths: list[Path],
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str | None = None,
    temperature: float = 0.2,
) -> dict[str, Any]:
    templates = scan_template_library()
    raw_texts = [extract_text(path, limit=None) for path in upload_paths]
    resolved_template_id = choose_template_id(prompt, raw_texts, template_id)
    template = next((item for item in templates if item["id"] == resolved_template_id), None)
    warnings: list[str] = []
    template_label = template["label"] if template else resolved_template_id
    materials = []
    for path in upload_paths:
        material, material_warnings = prepare_material(
            path=path,
            prompt=prompt,
            template_label=template_label,
            request_url=request_url,
            api_key=api_key,
            model_name=model_name,
            force_summary=resolved_template_id in SUMMARY_FIRST_TEMPLATE_IDS,
        )
        materials.append(material)
        warnings.extend(material_warnings)

    messages = build_ai_messages(prompt, materials, template)
    try:
        draft = call_chat_model(request_url, api_key, model_name, messages, temperature)
    except Exception as exc:
        draft = fallback_draft(prompt, materials, template["label"] if template else "公文")
        warnings.append(str(exc))

    raw_id = uuid.uuid4().hex
    raw_path = TEMP_DIR / f"{raw_id}_ai_raw.docx"
    result_id = uuid.uuid4().hex
    result_path = TEMP_DIR / f"{result_id}_ai.docx"
    write_plain_docx(draft, raw_path)
    convert_with_roles(str(raw_path), str(result_path), template_id=resolved_template_id)

    paragraphs = build_structure(str(result_path), resolved_template_id)
    filename = f"AI公文_{template['label'] if template else resolved_template_id}.docx"
    return {
        "documentId": result_id,
        "path": str(result_path),
        "filename": filename,
        "templateId": resolved_template_id,
        "templateName": template["name"] if template else "",
        "preview": draft[:3000],
            "paragraphs": paragraphs,
            "warnings": warnings,
            "readReport": materials[0].get("readReport") if materials else build_read_report({"items": [], "stats": _new_stats(), "warnings": []}),
        }


def generate_documents(
    prompt: str,
    upload_paths: list[Path],
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str | None = None,
    temperature: float = 0.2,
) -> dict[str, Any]:
    templates = scan_template_library()
    documents: list[dict[str, Any]] = []
    targets = [
        {"path": path, "name": display_name_from_temp_path(path), "text": extract_text(path, limit=None)}
        for path in upload_paths
    ]
    if not targets:
        targets = [{"path": None, "name": "new_document", "text": ""}]

    for index, material in enumerate(targets, start=1):
        resolved_template_id = choose_template_id(prompt, [material["text"]], template_id)
        template = next((item for item in templates if item["id"] == resolved_template_id), None)
        warnings: list[str] = []
        template_label = template["label"] if template else resolved_template_id
        if material["path"] is None:
            empty_report = build_read_report({"items": [], "stats": _new_stats(), "warnings": []})
            current_materials = [{
                "name": material["name"],
                "text": material["text"],
                "fullTextLength": len(material["text"]),
                "chunkCount": 0,
                "processingMode": "direct",
                "readReport": empty_report,
            }]
        else:
            prepared_material, material_warnings = prepare_material(
                path=material["path"],
                prompt=prompt,
                template_label=template_label,
                request_url=request_url,
                api_key=api_key,
                model_name=model_name,
                force_summary=resolved_template_id in SUMMARY_FIRST_TEMPLATE_IDS,
            )
            current_materials = [prepared_material]
            warnings.extend(material_warnings)
        messages = build_ai_messages(prompt, current_materials, template)

        try:
            draft = call_chat_model(request_url, api_key, model_name, messages, temperature)
        except Exception as exc:
            draft = fallback_draft(prompt, current_materials, template["label"] if template else "公文")
            warnings.append(str(exc))

        raw_id = uuid.uuid4().hex
        raw_path = TEMP_DIR / f"{raw_id}_ai_raw.docx"
        result_id = uuid.uuid4().hex
        result_path = TEMP_DIR / f"{result_id}_ai.docx"
        write_plain_docx(draft, raw_path)
        convert_with_roles(str(raw_path), str(result_path), template_id=resolved_template_id)

        paragraphs = build_structure(str(result_path), resolved_template_id)
        stem = Path(material["name"]).stem or f"document_{index}"
        documents.append({
            "documentId": result_id,
            "path": str(result_path),
            "filename": f"{stem}_AI公文.docx",
            "sourceName": material["name"],
            "templateId": resolved_template_id,
            "templateName": template["name"] if template else "",
            "preview": draft[:3000],
            "paragraphs": paragraphs,
            "warnings": warnings,
            "processingMode": current_materials[0].get("processingMode", "direct"),
            "fullTextLength": current_materials[0].get("fullTextLength", 0),
            "chunkCount": current_materials[0].get("chunkCount", 0),
            "readReport": current_materials[0].get("readReport"),
        })

    first = documents[0]
    return {
        **first,
        "documents": documents,
        "warnings": [warning for item in documents for warning in item.get("warnings", [])],
    }


def display_name_from_temp_path(path: Path) -> str:
    name = path.name
    if "_" not in name:
        return name
    _, encoded_name = name.split("_", 1)
    stem = Path(encoded_name).stem
    suffix = Path(encoded_name).suffix
    return f"{urllib.parse.unquote(stem)}{suffix}"
