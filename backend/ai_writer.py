"""AI document writing service helpers.

The implementation is intentionally OpenAI-compatible so it can point at an
intranet model gateway. Users provide request_url, api_key and model_name from
the assistant panel instead of relying on hard-coded cloud settings.
"""

from __future__ import annotations

import json
import math
import os
import re
import sqlite3
import uuid
import urllib.error
import urllib.parse
import urllib.request
from difflib import SequenceMatcher
from time import perf_counter
from collections import Counter
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from config import (
    get_document_config,
    get_template_definition,
    get_template_catalog,
    get_template_classification,
    normalize_template_id,
)
from converter import build_structure, convert_with_roles, normalize_role_for_text

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
READ_REPORT_SAMPLE_LIMIT = 20
# 期望成文字数下限（每份）；成文硬上限仍由 MAX_DRAFT_CHARS 控制。
MIN_DRAFT_CHARS = int(os.getenv("AI_MIN_DRAFT_CHARS", "5000"))
# 成文硬上限（每份）
MAX_DRAFT_CHARS = int(os.getenv("AI_MAX_DRAFT_CHARS", "30000"))
# 为避免“材料先归纳后成文”场景把上限误压到 5000，设置基础输出上限地板。
MIN_OUTPUT_LIMIT_CHARS = int(os.getenv("AI_MIN_OUTPUT_LIMIT_CHARS", "12000"))
# 单次模型 HTTP 调用超时（秒）。高并发排队时单次推理可能很久，默认放宽到 10 分钟。
MODEL_HTTP_TIMEOUT_SECONDS = int(os.getenv("AI_MODEL_HTTP_TIMEOUT_SECONDS", "600"))
# 按速度档配置输出 token 预算；5000 汉字约需 8k–12k tokens，标准档需留余量。
DRAFT_TOKEN_BUDGETS = {
    "fast": (8192,),
    "standard": (12288,),
    "deep": (24576,),
}
# 完整材料直送模型；超过此上限时拒绝请求，绝不静默截断或自动摘要。
MAX_DIRECT_INPUT_CHARS = int(os.getenv("AI_MAX_DIRECT_INPUT_CHARS", "100000"))
MAX_REFERENCE_INPUT_CHARS = int(os.getenv("AI_MAX_REFERENCE_INPUT_CHARS", "50000"))
MAX_GENERATION_GUARD_ATTEMPTS = int(os.getenv("AI_MAX_GENERATION_GUARD_ATTEMPTS", "5"))
SIMILARITY_REWRITE_COPY_THRESHOLD = float(os.getenv("AI_SIMILARITY_COPY_THRESHOLD", "0.90"))
SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD = float(os.getenv("AI_SIMILARITY_PARAGRAPH_THRESHOLD", "0.30"))
SIMILARITY_FAILURE_MARKERS = (
    "参考范文事实",
    "重合度过高",
    "重合度偏高",
    "复用参考范文",
    "直接复刻",
    "范文独有",
    "相似度偏高",
    "范文表述串入",
)

# 材料上传仅接受当前能可靠抽取的格式；wps/ofd 仅可入库为模板索引，需先转 DOCX/PDF。
SUPPORTED_INPUT_SUFFIXES = {".docx", ".txt", ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"}
UNREADABLE_REFERENCE_SUFFIXES = {".wps", ".ofd"}
REFERENCE_SUFFIXES = SUPPORTED_INPUT_SUFFIXES | UNREADABLE_REFERENCE_SUFFIXES
READABLE_REFERENCE_SUFFIXES = set(SUPPORTED_INPUT_SUFFIXES)
SUMMARY_FIRST_TEMPLATE_IDS = {"minutes"}
ALLOWED_DRAFT_ROLES = {
    "title", "subtitle", "recipient", "body", "h1", "h2", "h3", "h4",
    "attachment_head", "attachment_other", "sign_unit", "sign_date",
    "sign_contact", "security", "other",
}
ROLE_NAMES = {
    "title": "标题", "subtitle": "副标题", "recipient": "主送机关", "body": "正文",
    "h1": "一级标题", "h2": "二级标题", "h3": "三级标题", "h4": "四级标题",
    "attachment_head": "附件说明", "attachment_other": "其他附件", "sign_unit": "落款单位",
    "sign_date": "成文日期", "sign_contact": "联系信息", "security": "密级", "other": "其他",
}
PRESERVED_SOURCE_ROLES = {
    "title", "subtitle", "recipient", "attachment_head", "attachment_other",
    "sign_unit", "sign_date", "sign_contact", "security",
}
RE_ATTACHMENT_START = re.compile(r"^附件\s*[：:0-9]")
RE_DOC_TITLE = re.compile(
    r"^(关于.+的(?:函|通知|意见|报告|请示|通报|决定|公告|通告|纪要|方案|办法|规定|清单|说明|工作))"
    r"|^(公安工作提示函|.*[函通知意见报告请示通报]$)"
)
RE_RECIPIENT_LINE = re.compile(r"^[\u4e00-\u9fffA-Za-z0-9（）()·\s]{2,48}[：:]$")

RE_DEGENERATE_PLACEHOLDER_LINE = re.compile(r"^(?P<label>[^：:\n]{2,20}[：:])\s*(?:[、,，\s]{8,}|(?:\S?\s*[、,，]\s*){8,})$")
RE_PLACEHOLDER_TEXT = re.compile(
    r"(请补充|待补|待定|待填写|XXXX+|X{2,}|×{2,}|某单位|某某|"
    r"[\[【](?:[^\]】]{0,20})(?:主要领导|分管领导|处室名称|任务|目标|难点|方面|数量|时间)(?:[^\]】]{0,20})[\]】])",
    re.IGNORECASE,
)
RE_ROLE_LINE = re.compile(r"^\[\[(?P<role>[a-z0-9_]+)\]\]\s*(?P<text>.+)$")
RE_ROLE_MARKER = re.compile(r"^\[\[(?P<role>[a-z0-9_]+)\]\]\s*$")
RE_ALT_ROLE_LINE = re.compile(r"^\[(?P<role>[a-z0-9_]+)\]\s*(?P<text>.+)$")
RE_ALT_ROLE_MARKER = re.compile(r"^\[(?P<role>[a-z0-9_]+)\]\s*$")
RE_CN_ROLE_LINE = re.compile(r"^【(?P<role>[^】]+)】\s*(?P<text>.+)$")
RE_CN_ROLE_MARKER = re.compile(r"^【(?P<role>[^】]+)】\s*$")
RE_H1_TEXT = re.compile(r"^[一二三四五六七八九十百]+[、．.](?P<text>.+)$")
RE_H2_TEXT = re.compile(r"^[（(][一二三四五六七八九十百]+[）)](?P<text>.+)$")
RE_PAREN_SUBTITLE = re.compile(r"^[（(][^）)]{2,30}[）)]$")
ROLE_LABEL_ALIASES = {
    "标题": "title", "副标题": "subtitle", "主送机关": "recipient", "主送": "recipient",
    "正文": "body", "一级标题": "h1", "二级标题": "h2", "三级标题": "h3", "四级标题": "h4",
    "附件说明": "attachment_head", "其他附件": "attachment_other", "落款单位": "sign_unit",
    "成文日期": "sign_date", "联系信息": "sign_contact", "密级": "security", "其他": "other",
}
RE_EMPTY_SUMMARY = re.compile(r"(无有效材料|无可归纳内容|未提供.{0,8}材料)")
RE_EMPTY_MATERIAL_VALUE = re.compile(r"^(?:无|暂无|N/?A|-+)$", re.IGNORECASE)
RE_MATERIAL_METADATA = re.compile(
    r"(?:签发人|等级|发电时间|承办单位|发布日期|浏览次数|案号|案由|文号|发文字号)\s*[：:]?"
    r"|^[^\s]{1,24}〔\s*\d{4}\s*〕[^\s]{0,24}号$"
)
RE_REFERENCE_FACT_TOKEN = re.compile(
    r"\d+(?:\.\d+)+|\d{2,}|[〇一二三四五六七八九十]{4}年[〇一二三四五六七八九十月日]{1,8}"
)

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


def get_model_config(mask_secret: bool = False) -> dict[str, Any]:
    load_local_env()
    models = [
        item.strip()
        for item in os.getenv("AI_AVAILABLE_MODELS", "deepseek-v4-flash").split(",")
        if item.strip()
    ]
    current_model = os.getenv("DEEPSEEK_MODEL", models[0] if models else "deepseek-v4-flash")
    if current_model not in models:
        models.insert(0, current_model)
    api_key = os.getenv("DEEPSEEK_API_KEY", "")
    return {
        "requestUrl": os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
        "apiKey": ("*" * 8 if api_key and mask_secret else api_key),
        "apiKeyConfigured": bool(api_key),
        "modelName": current_model,
        "models": models,
    }


def save_model_config(request_url: str, api_key: str, model_name: str, models: list[str] | None = None) -> dict[str, Any]:
    clean_models = [item.strip() for item in (models or []) if item and item.strip()]
    if model_name.strip() and model_name.strip() not in clean_models:
        clean_models.insert(0, model_name.strip())
    if not clean_models:
        clean_models = ["deepseek-v4-flash"]

    load_local_env()
    existing_key = os.getenv("DEEPSEEK_API_KEY", "")
    incoming_key = api_key.strip()
    # 前端回传掩码时保留原密钥
    if not incoming_key or set(incoming_key) <= {"*"}:
        incoming_key = existing_key

    values = {
        "DEEPSEEK_BASE_URL": request_url.strip() or "https://api.deepseek.com",
        "DEEPSEEK_API_KEY": incoming_key,
        "DEEPSEEK_MODEL": model_name.strip() or clean_models[0],
        "AI_AVAILABLE_MODELS": ",".join(clean_models),
    }
    ENV_PATH.write_text("\n".join(f"{key}={value}" for key, value in values.items()) + "\n", encoding="utf-8")
    for key, value in values.items():
        os.environ[key] = value
    return get_model_config(mask_secret=True)


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


def _xml_text(element) -> str:
    return "\n".join(
        text
        for text in ("".join(t.text or "" for t in paragraph.findall(".//" + qn("w:t"))).strip()
                     for paragraph in element.findall(".//" + qn("w:p")))
        if text
    )


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
            seen_cells: set[Any] = set()
            try:
                rows = list(block.rows)
            except Exception as exc:
                warnings.append(f"{path.name} table {block_index} skipped structured table read: {exc}")
                table_text = _xml_text(block._tbl)
                if table_text:
                    _add_inventory_item(items, stats, "table", f"table {block_index} raw text", table_text)
                continue
            stats["tableRows"] += len(rows)
            for row_index, row in enumerate(rows, start=1):
                try:
                    cells = list(row.cells)
                except Exception as exc:
                    warnings.append(f"{path.name} table {block_index} row {row_index} skipped cell read: {exc}")
                    row_text = _xml_text(row._tr)
                    if row_text:
                        _add_inventory_item(
                            items, stats, "table",
                            f"table {block_index} row {row_index} raw text",
                            row_text,
                        )
                    continue
                for cell_index, cell in enumerate(cells, start=1):
                    # python-docx 会为横向合并单元格重复返回同一个底层 tc。
                    cell_key = cell._tc
                    if cell_key in seen_cells:
                        continue
                    seen_cells.add(cell_key)
                    stats["tableCells"] += 1
                    try:
                        cell_text = "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                    except Exception as exc:
                        warnings.append(
                            f"{path.name} table {block_index} row {row_index} cell {cell_index} skipped paragraph read: {exc}"
                        )
                        cell_text = _xml_text(cell._tc)
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


def _collect_ocr_inventory(path: Path) -> dict[str, Any]:
    """对图片及扫描 PDF 进行按需 OCR，依赖缺失时给出可操作的错误。"""
    try:
        import pytesseract
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError(
            "PDF/图片 OCR 依赖未安装（需要 pytesseract、Pillow）。"
            "请安装后端依赖，或改为上传可复制文本的 DOCX/TXT"
        ) from exc
    tesseract_cmd = os.getenv("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if Path(tesseract_cmd).exists():
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    tessdata_dir = os.getenv("TESSDATA_PREFIX", str(BASE_DIR / "ocr-tessdata"))
    if Path(tessdata_dir).exists():
        os.environ["TESSDATA_PREFIX"] = tessdata_dir
    try:
        available_languages = pytesseract.get_languages(config="")
    except Exception as exc:
        raise RuntimeError(
            f"无法启动本机 Tesseract OCR（{exc}）。"
            "请安装 Tesseract 并配置 TESSERACT_CMD，或改为上传 DOCX/TXT"
        ) from exc
    if "chi_sim" not in available_languages:
        raise RuntimeError(
            "未检测到 Tesseract 简体中文语言包 chi_sim，无法识别中文扫描件。"
            "请安装语言包，或改为上传可复制文本的 DOCX/TXT"
        )
    stats = _new_stats()
    items: list[dict[str, Any]] = []
    warnings: list[str] = []
    images = []
    if path.suffix.lower() == ".pdf":
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError(
                "PDF OCR 依赖未安装（需要 PyMuPDF/fitz）。"
                "请安装后端依赖，或先将 PDF 转为 DOCX/TXT 再上传"
            ) from exc
        pdf = fitz.open(path)
        for page in pdf:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
    else:
        images.append(Image.open(path))
    for index, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image, lang="chi_sim+eng").strip()
        if text:
            _add_inventory_item(items, stats, "ocr", f"page {index}", text)
    if not items:
        warnings.append("OCR 未识别到可用文字，请确认图像清晰度及本机中文 OCR 语言包")
    return {"items": items, "stats": stats, "warnings": warnings}


def extract_document_inventory(path: Path) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return _collect_docx_inventory(path)
    if suffix == ".txt":
        return _collect_txt_inventory(path)
    if suffix in {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"}:
        return _collect_ocr_inventory(path)

    stats = _new_stats()
    warning = f"Unsupported input file type: {suffix}. Convert it to docx, txt, PDF, or image first."
    return {"items": [], "stats": stats, "warnings": [warning]}


def inventory_to_text(inventory: dict[str, Any]) -> str:
    lines = []
    for item in inventory.get("items", []):
        lines.append(f"[{item['id']} | {item['sourceType']} | {item['location']}]\n{item['text']}")
    return "\n\n".join(lines).strip()


def assess_material_content(inventory: dict[str, Any]) -> dict[str, Any]:
    """区分可归纳的业务内容与文件控制元数据，防止元数据表被误生成为公文。"""
    substantive_chars = 0
    metadata_chars = 0
    substantive_items = 0
    for item in inventory.get("items", []):
        text = str(item.get("text") or "").strip()
        source_type = item.get("sourceType")
        if not text or RE_EMPTY_MATERIAL_VALUE.match(text):
            continue
        is_metadata = source_type in {"header", "footer"} or (
            source_type == "table" and bool(RE_MATERIAL_METADATA.search(text))
        )
        if is_metadata:
            metadata_chars += len(text)
        else:
            substantive_items += 1
            substantive_chars += len(text)
    if substantive_items:
        status = "ready"
        reason = ""
    elif metadata_chars:
        status = "metadata_only"
        reason = "文件仅提取到文号、签发人、等级、时间等元数据，未发现可用于归纳的正文或业务表格内容"
    else:
        status = "empty"
        reason = "文件未提取到可用于归纳的正文或业务表格内容"
    return {
        "status": status,
        "substantiveChars": substantive_chars,
        "substantiveItems": substantive_items,
        "metadataChars": metadata_chars,
        "reason": reason,
    }


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
        "materialQuality": assess_material_content(inventory),
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


def safe_extract_text(path: Path, limit: int | None = DIRECT_DRAFT_CHAR_LIMIT) -> tuple[str, list[str]]:
    try:
        inventory = extract_document_inventory(path)
        text = inventory_to_text(inventory).strip()
        if limit is not None and len(text) > limit:
            text = text[:limit] + "\n...[content truncated]"
        return text, list(inventory.get("warnings") or [])
    except Exception as exc:
        return "", [f"{display_name_from_temp_path(path)} text extraction failed: {exc}"]


def _source_ids_in_text(text: str) -> list[str]:
    ids = []
    for part in text.split("["):
        if "]" not in part:
            continue
        source_id = part.split("]", 1)[0].split("|", 1)[0].strip()
        if source_id and source_id not in ids:
            ids.append(source_id)
    return ids


def prepare_material(
    path: Path,
    prompt: str,
    template_label: str,
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str,
    inventory: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], list[str]]:
    name = display_name_from_temp_path(path)
    inventory = inventory or extract_document_inventory(path)
    raw_text = inventory_to_text(inventory)
    material_zones = analyze_material_zones(inventory, path, template_id)
    warnings: list[str] = list(inventory.get("warnings") or [])
    material_quality = assess_material_content(inventory)
    if material_quality["status"] != "ready":
        raise RuntimeError(f"材料正文不足：{material_quality['reason']}。请上传正文页、附件或完整文件")
    if len(raw_text) > MAX_DIRECT_INPUT_CHARS:
        raise RuntimeError(
            f"{name} 提取到 {len(raw_text)} 个字符，超过完整文本直送上限 {MAX_DIRECT_INPUT_CHARS}。"
            "为避免截断或自动摘要，已停止生成；请拆分材料或提高 AI_MAX_DIRECT_INPUT_CHARS。"
        )
    processed = process_upload_material_with_ai(
        inventory, prompt, template_label, template_id,
        request_url, api_key, model_name,
        material_zones=material_zones,
    )
    if processed["processingMode"] == "ai_structured_brief":
        warnings.append("已对上传材料进行 AI 归纳润色，后续依体例成文基于整理底稿而非原文直送")
    else:
        warnings.append("材料 AI 润色不可用，已回退为规则化事实清单")
    source_elements = extract_source_elements(path, template_id, inventory)
    brief = processed["briefForDrafting"]
    if material_zones.get("attachmentFactIds"):
        warnings.append(
            f"已识别附件区 {len(material_zones['attachmentFactIds'])} 段材料，"
            "成文时仅写入附件说明，不展开附件正文"
        )
    prepared_title = str(processed.get("title") or "").strip()
    if prepared_title and not looks_like_section_heading(prepared_title):
        material_zones["titleText"] = prepared_title
        if not any(item.get("role") == "title" for item in source_elements):
            source_elements.insert(0, {"role": "title", "text": prepared_title})
        else:
            for item in source_elements:
                if item.get("role") == "title":
                    item["text"] = prepared_title
                    break
    return {
        "name": name,
        "text": brief,
        "rawText": processed.get("rawText") or raw_text,
        "summary": processed.get("summary") or "",
        "title": prepared_title,
        "structuredFacts": processed.get("facts") or [],
        "attachmentBrief": processed.get("attachmentBrief") or "",
        "fullTextLength": len(brief),
        "rawTextLength": len(raw_text),
        "chunkCount": 1 if brief else 0,
        "processingMode": processed["processingMode"],
        "readReport": build_read_report(inventory),
        "materialQuality": material_quality,
        "materialZones": material_zones,
        "sourceElements": source_elements,
    }, warnings


def draft_char_limit(materials: list[dict[str, Any]], reference: dict[str, Any]) -> int:
    """按材料与体例估算成文上限；避免归纳底稿过短导致上限被误压。"""
    source_chars = 0
    for item in materials:
        brief_chars = int(item.get("fullTextLength") or len(item.get("text", "")))
        raw_chars = int(item.get("rawTextLength") or len(item.get("rawText", "")))
        # 归纳底稿用于质量，原文长度用于上限估算，避免“超过5000字上限”误判。
        source_chars += max(brief_chars, int(raw_chars * 0.9))
    reference_chars = int(reference.get("textChars") or len(reference.get("text", "")))
    return min(
        MAX_DRAFT_CHARS,
        max(
            MIN_OUTPUT_LIMIT_CHARS,
            int(max(source_chars, MIN_DRAFT_CHARS) * 1.35),
            int(reference_chars * 1.15),
        ),
    )


def analyze_material_zones(
    inventory: dict[str, Any],
    path: Path | None = None,
    template_id: str = "generic",
) -> dict[str, Any]:
    """识别上传材料中的主件（函/通知正文）与附件区，供成文时分区使用。"""
    zones: dict[str, Any] = {
        "titleText": "",
        "recipientText": "",
        "attachmentHeadText": "",
        "signUnits": [],
        "signDateText": "",
        "signContactText": "",
        "letterFactIds": [],
        "attachmentFactIds": [],
        "signatureFactIds": [],
        "summary": "",
    }
    structure_items: list[dict[str, Any]] = []
    if path and path.suffix.lower() == ".docx":
        try:
            structure_items = build_structure(str(path), template_id)
        except Exception:
            structure_items = []

    if structure_items:
        zone = "letter"
        seen_signature = False
        title_lines: list[str] = []
        for item in structure_items:
            role = str(item.get("role") or "body")
            text = str(item.get("fullText") or "").strip()
            if not text:
                continue
            if role == "title":
                title_lines.append(text)
                continue
            if title_lines and role not in {"subtitle", "recipient"}:
                zones["titleText"] = "\n".join(title_lines)
                title_lines = []
            if role == "subtitle" and not zones["titleText"]:
                zones["titleText"] = text
                continue
            if role == "recipient":
                zones["recipientText"] = text
                zone = "letter"
                continue
            if role == "attachment_head":
                zones["attachmentHeadText"] = text
                zone = "attachment"
                continue
            if role in {"sign_unit", "sign_date", "sign_contact"}:
                seen_signature = True
                zone = "signature"
                if role == "sign_unit":
                    zones["signUnits"].append(text)
                elif role == "sign_date":
                    zones["signDateText"] = text
                else:
                    zones["signContactText"] = text
                continue
            if seen_signature and role in {"title", "h1", "h2", "body"}:
                zone = "attachment"
            elif role == "attachment_other":
                zone = "attachment"
        if title_lines:
            zones["titleText"] = "\n".join(title_lines)
        # 结构识别偶发把「一、…」标成 title：清空，留给主标题提炼
        if looks_like_section_heading(zones.get("titleText") or ""):
            zones["titleText"] = ""

    items = [item for item in inventory.get("items") or [] if item.get("sourceType") not in {"header", "footer"}]
    state = {"after_signature": False, "body_seen": False}
    for item in items:
        fact_id = str(item.get("id") or "")
        text = str(item.get("text") or "").strip()
        if not text or not fact_id:
            continue
        if (
            not zones["titleText"]
            and RE_DOC_TITLE.match(text)
            and not looks_like_section_heading(text)
        ):
            zones["titleText"] = text
        elif (
            not zones["titleText"]
            and 6 <= len(text) <= 40
            and not looks_like_section_heading(text)
            and not RE_RECIPIENT_LINE.match(text)
            and ("会议" in text or "材料" in text or "调度" in text or text.endswith(("报告", "汇报", "情况")))
        ):
            zones["titleText"] = text
        if not zones["recipientText"] and RE_RECIPIENT_LINE.match(text):
            zones["recipientText"] = text
        if RE_ATTACHMENT_START.match(text):
            zones["attachmentHeadText"] = zones["attachmentHeadText"] or text
        if re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", text):
            zones["signDateText"] = zones["signDateText"] or text
            state["after_signature"] = True
        elif (
            re.search(r"(?:厅|局|部|委|办|院|队|中心|处|组|总队|支队|大队|政府|委员会)$", text)
            and 4 <= len(text) <= 40
            and text not in zones["titleText"]
            and not RE_RECIPIENT_LINE.match(text)
            and state["body_seen"]
        ):
            if text not in zones["signUnits"]:
                zones["signUnits"].append(text)

        if state["after_signature"]:
            zones["attachmentFactIds"].append(fact_id)
        elif (
            RE_DOC_TITLE.match(text)
            or RE_RECIPIENT_LINE.match(text)
            or RE_ATTACHMENT_START.match(text)
            or re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", text)
            or (
                text in zones["signUnits"]
                and re.search(r"(?:厅|局|部|委|办|院|队|中心|处|组|总队|支队|大队|政府|委员会)$", text)
            )
        ):
            zones["signatureFactIds"].append(fact_id)
        else:
            zones["letterFactIds"].append(fact_id)
            state["body_seen"] = True

    if zones["titleText"] or zones["attachmentHeadText"] or zones["signUnits"]:
        parts = []
        if zones["titleText"]:
            parts.append(f"主件标题：{zones['titleText']}")
        if zones["recipientText"]:
            parts.append(f"主送：{zones['recipientText']}")
        if zones["attachmentHeadText"]:
            parts.append(f"附件说明：{zones['attachmentHeadText']}")
        if zones["signUnits"]:
            parts.append(f"落款：{' / '.join(zones['signUnits'])}")
        zones["summary"] = "；".join(parts)
    return zones


def looks_like_section_heading(text: str) -> bool:
    """判断是否为「一、」「（一）」类章节小标题，而非公文主标题。"""
    value = (text or "").strip()
    if not value:
        return False
    if RE_H1_TEXT.match(value) or RE_H2_TEXT.match(value):
        return True
    return False


def looks_like_parenthetical_subtitle(text: str) -> bool:
    return bool(RE_PAREN_SUBTITLE.match((text or "").strip()))


def render_material_zones_text(zones: dict[str, Any]) -> str:
    """生成正向材料分区说明，帮助模型理解主件与附件边界。"""
    lines = [
        "上传材料结构识别（成文时按此分区取材）：",
        f"- 主件正文区（写入函/通知主体）：{len(zones.get('letterFactIds') or [])} 段事实",
        "- 成文完整性：独立主标题 [[title]]、主体正文、材料具备时的落款；"
        "「一、二、三…」只作 [[h1]]，不作主标题。",
    ]
    if zones.get("titleText") and not looks_like_section_heading(zones["titleText"]):
        lines.append(f"- 主件标题（独立 [[title]] 段）：{zones['titleText']}")
    elif not zones.get("titleText") or looks_like_section_heading(zones.get("titleText") or ""):
        lines.append("- 主件标题：材料中需从会议/主题名称提炼独立 [[title]]（勿用一级小标题充当）")
    if zones.get("recipientText"):
        lines.append(f"- 主送机关（独立 [[recipient]] 段）：{zones['recipientText']}")
    if zones.get("attachmentHeadText"):
        lines.append(f"- 附件说明（仅一行 [[attachment_head]]）：{zones['attachmentHeadText']}")
    if zones.get("attachmentFactIds"):
        lines.append(
            f"- 附件正文区（{len(zones['attachmentFactIds'])} 段，供了解背景；"
            "不写入函件主体，仅在附件说明中点名即可）"
        )
    if zones.get("signUnits"):
        lines.append(f"- 落款单位（每行一个 [[sign_unit]]）：" + "｜".join(zones["signUnits"]))
    if zones.get("signDateText"):
        lines.append(f"- 成文日期（[[sign_date]]）：{zones['signDateText']}")
    if zones.get("signContactText"):
        lines.append(f"- 联系信息（[[sign_contact]]）：{zones['signContactText']}")
    return "\n".join(lines)


def extract_source_elements(
    path: Path,
    template_id: str,
    inventory: dict[str, Any] | None = None,
) -> list[dict[str, str]]:
    """提取上传材料中已有的公文要素，作为生成时不可丢失的内容。"""
    zones = analyze_material_zones(inventory or {"items": []}, path, template_id)
    if path.suffix.lower() != ".docx":
        elements: list[dict[str, str]] = []
        if zones.get("titleText"):
            elements.append({"role": "title", "text": zones["titleText"]})
        if zones.get("recipientText"):
            elements.append({"role": "recipient", "text": zones["recipientText"]})
        if zones.get("attachmentHeadText"):
            elements.append({"role": "attachment_head", "text": zones["attachmentHeadText"]})
        for unit in zones.get("signUnits") or []:
            elements.append({"role": "sign_unit", "text": unit})
        if zones.get("signDateText"):
            elements.append({"role": "sign_date", "text": zones["signDateText"]})
        if zones.get("signContactText"):
            elements.append({"role": "sign_contact", "text": zones["signContactText"]})
        return elements
    try:
        items = build_structure(str(path), template_id)
    except Exception:
        return []

    elements: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    title_parts: list[str] = []
    header_phase = True
    for item in items:
        role = item.get("role")
        text = str(item.get("fullText") or "").strip()
        if not text:
            continue
        if role == "title":
            # 「一、…」是章节标题，不作为可保留主标题
            if looks_like_section_heading(text):
                continue
            title_parts.append(text)
            continue
        if title_parts:
            merged = "\n".join(title_parts)
            if not looks_like_section_heading(merged):
                key = ("title", merged)
                if key not in seen:
                    elements.append({"role": "title", "text": merged})
                    seen.add(key)
            title_parts = []
        if role == "subtitle" or (
            header_phase
            and looks_like_parenthetical_subtitle(text)
            and any(item["role"] == "title" for item in elements)
        ):
            key = ("subtitle", text)
            if key not in seen:
                elements.append({"role": "subtitle", "text": text})
                seen.add(key)
            continue
        if role in {"h1", "h2", "h3", "h4", "body"}:
            header_phase = False
        if role not in PRESERVED_SOURCE_ROLES:
            continue
        key = (role, text)
        if key not in seen:
            elements.append({"role": role, "text": text})
            seen.add(key)
    if title_parts:
        merged = "\n".join(title_parts)
        if not looks_like_section_heading(merged):
            key = ("title", merged)
            if key not in seen:
                elements.append({"role": "title", "text": merged})
    return elements


def guess_template_id(text: str, filename: str = "") -> str:
    """结合文件夹/文件名前缀与正文关键词打分识别文种。"""
    from collections import Counter
    scores: Counter[str] = Counter()
    combined = "\n".join(part for part in (filename, text) if part)
    for template_id, keywords in DOCUMENT_TYPE_HINTS:
        for keyword in keywords:
            if keyword in combined:
                scores[template_id] += 3 if keyword in filename else 1
            if filename and re.search(rf"\d{{1,2}}\s*{re.escape(keyword)}", filename):
                scores[template_id] += 5
    if not scores:
        return "generic"
    ranked = sorted(scores.items(), key=lambda item: (-item[1], item[0]))
    if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
        return "generic"
    return ranked[0][0]


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
                corpus_profile_json TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL,
                PRIMARY KEY (id, source_dir)
            )
            """
        )
        columns = {row[1] for row in conn.execute("PRAGMA table_info(document_templates)")}
        if "corpus_profile_json" not in columns:
            conn.execute(
                "ALTER TABLE document_templates ADD COLUMN corpus_profile_json TEXT NOT NULL DEFAULT '{}'"
            )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS template_references (
                template_id TEXT NOT NULL,
                source_dir TEXT NOT NULL,
                filename TEXT NOT NULL,
                file_path TEXT NOT NULL,
                suffix TEXT NOT NULL,
                status TEXT NOT NULL,
                warning TEXT NOT NULL DEFAULT '',
                text TEXT NOT NULL DEFAULT '',
                text_chars INTEGER NOT NULL DEFAULT 0,
                roles_json TEXT NOT NULL DEFAULT '[]',
                structure_json TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL,
                PRIMARY KEY (template_id, source_dir, filename)
            )
            """
        )
        ref_columns = {row[1] for row in conn.execute("PRAGMA table_info(template_references)")}
        if "structure_json" not in ref_columns:
            conn.execute(
                "ALTER TABLE template_references ADD COLUMN structure_json TEXT NOT NULL DEFAULT '{}'"
            )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_document_templates_id ON document_templates(id)")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_template_references_lookup "
            "ON template_references(template_id, source_dir, status)"
        )


def load_templates_from_db() -> list[dict[str, Any]]:
    init_template_db()
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT id, name, label, file_count, source_dir, sample_file,
                   sample_path, sample_text, files_json, corpus_profile_json, updated_at
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
        try:
            corpus_profile = json.loads(row["corpus_profile_json"] or "{}")
        except json.JSONDecodeError:
            corpus_profile = {}
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
            "corpusProfile": corpus_profile,
            "analyzedFileCount": int(corpus_profile.get("analyzedFileCount", 0)),
            "availableReferenceCount": int(corpus_profile.get("availableReferenceCount", 0)),
            "unreadableReferenceCount": len(corpus_profile.get("skippedFiles") or []),
            "corpusStatus": corpus_profile.get("status", "empty"),
            "updatedAt": row["updated_at"],
        })
    return sorted(templates, key=_template_sort_key)


def _compact_roles(roles: list[str]) -> list[str]:
    compact: list[str] = []
    for role in roles:
        if role != "other" and (not compact or compact[-1] != role):
            compact.append(role)
    return compact


def _default_section_function(role: str) -> str:
    defaults = {
        "title": "概括公文主题",
        "subtitle": "补充说明主题",
        "recipient": "明确主送机关",
        "body": "阐述背景、依据、事项或要求",
        "h1": "划分主体一级层次",
        "h2": "划分主体二级层次",
        "h3": "划分主体三级层次",
        "h4": "划分主体四级层次",
        "attachment_head": "说明附件",
        "attachment_other": "列示其他附件",
        "sign_unit": "落款单位",
        "sign_date": "成文日期",
        "sign_contact": "联系信息",
        "security": "密级标识",
    }
    return defaults.get(role, "按文种常规职责组织")


def _default_writing_method(role: str) -> str:
    defaults = {
        "title": "从材料主题中提炼简明标题，居中、无句号",
        "body": "用本段对应材料事实完成叙述：先写依据与背景，再写事项与要求，必要时以一是二是分项展开",
        "h1": "用「一、二、三…」划分主体大段，标题短语概括该段材料主题",
        "h2": "在大段内用「（一）（二）…」细分，序号随所属大段从（一）重计",
        "h3": "概括细分材料要点主题",
        "h4": "概括细分材料要点主题",
        "recipient": "材料已有主送机关时顶格写出并加冒号",
        "sign_unit": "材料已有落款单位时右对齐写出",
        "sign_date": "材料已有成文日期时右对齐、用汉字日期写出",
    }
    return defaults.get(role, "按本段职责，用对应材料事实完成公文语体表达")


def _default_format_rules(role: str) -> list[str]:
    rules = {
        "title": ["不超过30字", "不加标点"],
        "h1": ["小标题后另起正文段", "不以句号结尾"],
        "h2": ["小标题后另起正文段", "不以句号结尾"],
        "body": ["首行缩进2字符", "公文语体"],
    }
    return rules.get(role, [])


def _default_constraint_tags(role: str) -> list[str]:
    tags = {
        "title": ["derive_from_source", "no_new_facts"],
        "body": ["paraphrase_required", "no_verbatim_source"],
        "recipient": ["preserve_if_present"],
        "sign_unit": ["preserve_if_present"],
        "sign_date": ["preserve_if_present"],
    }
    return tags.get(role, ["no_new_facts"])


def _infer_organization(roles: list[str], template_id: str = "") -> dict[str, Any]:
    profile = build_genre_writing_profile(template_id) if template_id else {}
    if profile.get("organization"):
        return dict(profile["organization"])
    body_roles = [role for role in roles if role in {"body", "h1", "h2", "h3", "h4"}]
    opening = "标题、主送机关后进入缘由或依据段" if "recipient" in roles else "标题后进入缘由或依据段"
    development = "按标题层级展开主体内容" if any(r.startswith("h") for r in body_roles) else "主体段分层展开"
    closing = "末段提出工作要求或说明事项"
    if "sign_unit" in roles or "sign_date" in roles:
        closing = "主体后自然收束，最后落款"
    return {"opening": opening, "development": development, "closing": closing, "style": ["公文语体", "先总后分"]}


def _infer_style_tags(text: str, roles: list[str], template_id: str = "") -> list[str]:
    profile = build_genre_writing_profile(template_id) if template_id else {}
    if profile.get("mode") == "layout_sections":
        tags = ["版式分区", "连续正文"]
        if "recipient" in roles:
            tags.append("主送明确")
        return tags
    tags = []
    if any(role in roles for role in {"h1", "h2", "h3", "h4"}):
        tags.append("层级标题")
    if re.search(r"[一二三四五六七八九十]是", text):
        tags.append("分条叙述")
    if "recipient" in roles:
        tags.append("有主送")
    return tags or ["常规公文"]


def _infer_rhetoric_patterns(text: str, roles: list[str], template_id: str = "") -> list[str]:
    profile = build_genre_writing_profile(template_id) if template_id else {}
    if profile.get("mode") == "layout_sections":
        patterns = ["公文语体", "前缀-正文-落款", "事由与事项连贯叙述"]
        if "recipient" in roles:
            patterns.append("主送后直接展开事项")
        return patterns
    patterns = ["公文语体", "先述依据后述事项"]
    if "recipient" in roles:
        patterns.append("主送后展开缘由")
    if any(role in roles for role in {"h1", "h2"}):
        patterns.append("分项标题统领正文")
    if re.search(r"特此|现将|现予|请认真贯彻执行", text):
        patterns.append("结尾固定收束语可选")
    if re.search(r"[一二三四五六七八九十][、．.]", text):
        patterns.append("序数分条叙述")
    return patterns


def _infer_scene_rules(template_id: str, roles: list[str]) -> list[str]:
    profile = build_genre_writing_profile(template_id)
    rules = [
        f"按「{profile.get('label') or template_id}」文种写作品格组织：{profile.get('purpose', '')}",
        "标题、主送与落款优先按上传材料要素成段",
    ]
    if profile.get("mode") == "layout_sections":
        rules.append("正文以连续段落完成商洽/告知事项，落款分区完整")
    elif profile.get("mode") == "layered_sections":
        rules.append("主体可按材料分块，标题后紧跟正文")
    if template_id == "minutes":
        rules.append("纪要按会议议题归纳，突出议定事项")
    if "sign_unit" in roles or "sign_date" in roles:
        rules.append("落款单位可多行输出，日期与联系方式独立成段")
    return rules


def _abstract_pattern_hint(snippet: str) -> str:
    """将范文片段抽象为写法提示，入库仅保留体例线索。"""
    text = re.sub(r"\d+(?:\.\d+)?", "N", snippet or "")
    text = re.sub(r"[〇一二三四五六七八九十]+年[〇一二三四五六七八九十月日]*", "日期", text)
    text = re.sub(r"[\u4e00-\u9fff]{2,12}(市|县|区|局|院|厅|办|部|委|队)", "某单位", text)
    text = re.sub(r"[\u4e00-\u9fff]{2,4}(某|某某)", "某人", text)
    return text[:48]


def render_structure_text(blueprint: dict[str, Any]) -> str:
    lines = [
        f"文种：{blueprint.get('templateId', '')}",
        f"段落序列：{' -> '.join(ROLE_NAMES.get(r, r) for r in blueprint.get('roleSequence', []))}",
        f"必备角色：{'、'.join(ROLE_NAMES.get(r, r) for r in blueprint.get('requiredRoles', [])) or '标题'}",
    ]
    org = blueprint.get("organization") or {}
    if org:
        lines.append(
            f"组织方式：开头={org.get('opening', '')}；主体={org.get('development', '')}；结尾={org.get('closing', '')}"
        )
    for pattern in blueprint.get("rhetoricPatterns") or []:
        lines.append(f"行文范式：{pattern}")
    for rule in blueprint.get("sceneRules") or []:
        lines.append(f"场景规则：{rule}")
    for section in blueprint.get("sections") or []:
        lines.append(
            f"{section.get('order')}. [{section.get('label')}] "
            f"职责={section.get('function')}；写法={section.get('writingMethod')}；"
            f"约束={','.join(section.get('constraintTags') or [])}"
        )
    return "\n".join(lines)


def build_rule_based_structure_blueprint(
    template_id: str,
    filename: str,
    roles: list[str],
    text: str,
    path: Path | None = None,
) -> dict[str, Any]:
    role_patterns: dict[str, str] = {}
    if path and path.suffix.lower() == ".docx":
        try:
            for item in build_structure(str(path), template_id):
                role = item.get("role")
                snippet = str(item.get("fullText") or "").strip()
                if role and snippet and role not in role_patterns:
                    role_patterns[role] = _abstract_pattern_hint(snippet)
        except Exception:
            pass
    compact_roles = _compact_roles(roles)
    if "title" not in compact_roles:
        compact_roles = ["title", *compact_roles]
    sections = []
    for order, role in enumerate(compact_roles, start=1):
        sections.append({
            "order": order,
            "role": role,
            "label": ROLE_NAMES.get(role, role),
            "function": _default_section_function(role),
            "writingMethod": _default_writing_method(role),
            "formatRules": _default_format_rules(role),
            "constraintTags": _default_constraint_tags(role),
            "paragraphFunction": _default_section_function(role),
            "patternHint": role_patterns.get(role, ""),
            "callable": True,
            "forbidden": ["verbatim_reference_text", "reference_facts"],
        })
    required_roles = ["title"]
    if "recipient" in compact_roles:
        required_roles.append("recipient")
    blueprint = {
        "version": 2,
        "storageFormat": "json",
        "templateId": template_id,
        "filename": filename,
        "sections": sections,
        "headingSkeleton": [role for role in compact_roles if role in {"h1", "h2", "h3", "h4"}],
        "roleSequence": compact_roles,
        "requiredRoles": required_roles,
        "organization": _infer_organization(compact_roles, template_id),
        "styleTags": _infer_style_tags(text, compact_roles, template_id),
        "rhetoricPatterns": _infer_rhetoric_patterns(text, compact_roles, template_id),
        "sceneRules": _infer_scene_rules(template_id, compact_roles),
        "callableFields": [
            "sections", "headingSkeleton", "roleSequence", "requiredRoles",
            "organization", "styleTags", "rhetoricPatterns", "sceneRules", "structureText",
            "genreGoal", "formatProfile", "featureVector", "writingProfile",
        ],
        "forbiddenFields": ["text", "sampleText", "verbatimBody", "referenceFacts"],
        "antiPlagiarism": {
            "reuseStructure": True,
            "reuseRhetoric": True,
            "reuseVerbatim": False,
            "rule": "以结构化段落职责与写法为体例坐标，用材料事实完成各段独立成文表达",
        },
    }
    # 预解析时固化版式与体裁目标，运行阶段直接读取缓存。
    resolved_config = get_document_config(template_id)
    classification = get_template_classification(template_id)
    blueprint["genreGoal"] = {
        "label": classification.get("label") or get_template_definition_label(template_id),
        "purpose": _infer_genre_goal(template_id),
        "requiredRecipient": classification.get("recipient", "optional"),
        "requiredSignature": classification.get("signature", "optional"),
        "forbiddenDirections": _infer_forbidden_directions(template_id),
        "writingProfile": build_genre_writing_profile(template_id),
    }
    if blueprint["genreGoal"]["writingProfile"].get("mode") == "layout_sections":
        blueprint["headingSkeleton"] = []
        blueprint["organization"] = blueprint["genreGoal"]["writingProfile"]["organization"]
        blueprint["writingProfile"] = blueprint["genreGoal"]["writingProfile"]
    else:
        blueprint["writingProfile"] = blueprint["genreGoal"]["writingProfile"]
    blueprint["formatProfile"] = {
        "page": resolved_config.get("page", {}),
        "spacing": resolved_config.get("spacing", {}),
        "styles": resolved_config.get("styles", {}),
        "blankLines": resolved_config.get("blank_lines", {}),
        "titleLevels": [section for section in sections if section.get("role") in {"title", "h1", "h2", "h3", "h4"}],
        "signatureRoles": [section for section in sections if section.get("role", "").startswith("sign_")],
        "sourceObservations": _extract_source_format_profile(path, template_id) if path else {},
    }
    blueprint["featureVector"] = _build_template_feature_vector(template_id, text, compact_roles, blueprint)
    blueprint["structureText"] = render_structure_text(blueprint)
    return blueprint


def get_template_definition_label(template_id: str) -> str:
    return str(get_template_definition(template_id).get("label") or template_id)


def _infer_genre_goal(template_id: str) -> str:
    goals = {
        "report": "汇报工作、反映情况、总结复盘，突出情况与结果",
        "request": "向上级请求指示或批准，突出请示事项与请求依据",
        "notice": "发布、传达事项或要求执行，突出对象、事项和执行要求",
        "letter": "不相隶属机关间商洽、询问、答复，突出事由清楚、事项明确、落款完整",
        "opinion": "提出见解、办法和工作要求，突出政策导向与实施安排",
        "minutes": "记载会议情况和议定事项，突出会议过程与决定结果",
        "announcement": "向社会公开重要事项，突出周知性与规范性",
        "public_notice": "在一定范围内公布应遵守或周知事项",
        "reply": "答复请示事项，突出答复意见与执行要求",
        "bulletin": "表彰、批评或传达重要情况，突出事由与要求",
        "decision": "作出重要决策或部署，突出决定事项",
        "resolution": "会议审议通过事项，突出决议内容",
        "order": "发布命令或任免授予事项，突出令文简洁明确",
        "proposal": "提请会议审议事项，突出议案主旨",
        "communique": "公开发布重要决定或情况，突出权威表述",
    }
    return goals.get(template_id, "围绕目标文种规范组织正式公文")


def _infer_forbidden_directions(template_id: str) -> list[str]:
    if template_id == "report":
        return ["请示", "申请", "请求批准"]
    if template_id == "notice":
        return ["个人申请", "请示式结尾"]
    return []


def build_genre_writing_profile(template_id: str) -> dict[str, Any]:
    """按文种给出正向写作品格：版式节奏、常用角色、成文示例，供策划与成文提示使用。"""
    classification = get_template_classification(template_id)
    label = classification.get("label") or get_template_definition_label(template_id)
    purpose = _infer_genre_goal(template_id)
    recipient = classification.get("recipient", "optional")
    signature = classification.get("signature", "optional")
    closing = classification.get("closing_phrases") or []

    # 扁平文种：函、公告、通告、命令等，以版式分区为主，正文连续成段
    flat_ids = {"letter", "announcement", "public_notice", "order", "communique", "reply"}
    # 层级文种：意见、通知、决定、报告等，可按材料设一二级标题
    layered_ids = {
        "opinion", "notice", "decision", "report", "request", "bulletin",
        "resolution", "proposal", "minutes",
    }

    if template_id in flat_ids or (template_id == "letter"):
        mode = "layout_sections"
        preferred_roles = ["title", "body", "attachment_head", "sign_unit", "sign_date", "sign_contact"]
        if recipient == "required":
            preferred_roles.insert(1, "recipient")
        elif recipient == "optional":
            preferred_roles.insert(1, "recipient")
        chapter_guide = (
            f"本篇按「{label}」文种写。函类/平行文注重版式分区："
            "前缀（标题、主送）→ 正文（通常一两段连贯叙述）→ 落款（单位可多行、日期、联系人及电话）。"
            "正文用 [[body]] 连续成段表达事由与事项；事项较多时用「一是、二是」写在正文段内。"
        )
        plan_guide = (
            "按版式分区设计 targetSections：[[title]]→[[recipient]]→一两段 [[body]]；"
            "如有附件则 [[attachment_head]] 一行；落款 [[sign_unit]] 每行一段、[[sign_date]]、[[sign_contact]]。"
            "本篇以正文段落完成叙述，结构蓝图聚焦 title/recipient/body/落款。"
        )
        example = (
            "[[title]]关于×××的函\n"
            "[[recipient]]××局：\n"
            "[[body]]根据工作安排，现就××事项函告如下：……（缘由与事项连贯叙述）\n"
            "[[body]]请贵单位予以协助，并将有关情况及时反馈我单位。\n"
            "[[attachment_head]]附件：××清单\n"
            "[[sign_unit]]××公安厅\n"
            "[[sign_unit]]××总队\n"
            "[[sign_date]]××××年×月×日\n"
            "[[sign_contact]]（联系人及电话：××，××××）\n"
        )
        heading_hint = "本篇文种以连续正文为主，版式重点在标题、主送、正文段与落款分区。"
        development = "正文一两段写清事由、事项与希望配合内容；分项可写在正文段内"
    elif template_id in layered_ids:
        mode = "layered_sections"
        preferred_roles = ["title", "body", "h1", "h2", "attachment_head", "sign_unit", "sign_date"]
        if recipient == "required":
            preferred_roles.insert(1, "recipient")
        # 会议材料等常有括注单位行
        if "subtitle" not in preferred_roles:
            preferred_roles.insert(1, "subtitle")
        chapter_guide = (
            f"本篇按「{label}」文种写。主体可按材料事实分块："
            "先写独立主标题 [[title]]（整篇主题，如会议名称），"
            "单位括注可用 [[subtitle]]；"
            "一级 [[h1]] 用「一、二、三…」划分大段，其下用 [[body]] 写清该段事实；"
            "内容较多时大段内用 [[h2]]「（一）（二）…」，二级序号在每个一级大段内从（一）重计。"
            "每个标题后紧跟对应正文。文末按材料写落款。"
        )
        plan_guide = (
            "按材料设计：[[title]]→（subtitle）→（主送）→导语 [[body]]→按需 h1→body 或 h1→h2→body；"
            "附件 [[attachment_head]] 一行；落款逐行 [[sign_unit]]、[[sign_date]]。"
            "targetSections 必须包含 title；「一、…」对应 h1 而非 title。"
        )
        example = (
            f"[[title]]×月×日××会有关材料\n"
            "[[subtitle]]（××总队）\n"
            "[[h1]]一、××工作推进情况\n"
            "[[body]]……本段正文……\n"
            "[[h1]]二、××下步安排\n"
            "[[body]]……本段正文……\n"
            "[[sign_unit]]××单位\n"
            "[[sign_date]]××××年×月×日\n"
        )
        heading_hint = "可按材料复杂度设置一级/二级标题；标题措辞由材料归纳。"
        development = "按事项逻辑分块展开，标题统领正文"
    else:
        mode = "adaptive"
        preferred_roles = ["title", "recipient", "body", "h1", "attachment_head", "sign_unit", "sign_date"]
        chapter_guide = (
            f"本篇按「{label}」文种写。先判断材料更适合连续正文还是分块标题："
            "事项单一时用一两段 [[body]]；事项较多时再用 [[h1]]/[[body]] 分块。"
        )
        plan_guide = (
            "按材料实际设计段落：先 title、（主送）、body；材料块多时再规划 h1→body。"
        )
        example = (
            f"[[title]]关于×××的{label}\n"
            "[[body]]……正文……\n"
            "[[sign_unit]]××单位\n"
            "[[sign_date]]××××年×月×日\n"
        )
        heading_hint = "按材料复杂度决定是否分块；内容少时连续正文即可。"
        development = "按文种常规与材料事实组织主体"

    opening_parts = ["输出独立 [[title]]"]
    if recipient != "forbidden":
        opening_parts.append("材料有主送则输出 [[recipient]]")
    opening_parts.append("再进入正文")

    closing_parts = []
    if closing:
        closing_parts.append(f"收束可自然使用文种惯用语（如{'、'.join(closing[:3])}）")
    if signature != "forbidden":
        closing_parts.append("落款单位可多行 [[sign_unit]]，日期 [[sign_date]]，联系人电话 [[sign_contact]]")
    else:
        closing_parts.append("按材料决定是否落款")

    return {
        "templateId": template_id,
        "label": label,
        "purpose": purpose,
        "mode": mode,
        "preferredRoles": preferred_roles,
        "chapterGuide": chapter_guide,
        "planGuide": plan_guide,
        "headingHint": heading_hint,
        "example": example,
        "organization": {
            "opening": "，".join(opening_parts),
            "development": development,
            "closing": "；".join(closing_parts),
            "style": ["公文语体", "版式清晰" if mode == "layout_sections" else "层次分明"],
        },
        "recipientPolicy": recipient,
        "signaturePolicy": signature,
    }


def _build_template_feature_vector(template_id: str, text: str, roles: list[str], blueprint: dict[str, Any]) -> dict[str, Any]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    terms = [token for token in re.findall(r"[\u4e00-\u9fff]{2,8}", text) if len(token) >= 2]
    return {
        "textChars": len(text), "paragraphCount": len(lines), "roleSequence": roles,
        "headingCount": sum(role in {"h1", "h2", "h3", "h4"} for role in roles),
        "titleLevelCount": len(blueprint.get("headingSkeleton") or []),
        "firstLine": lines[0][:80] if lines else "", "lastLine": lines[-1][:80] if lines else "",
        "styleTags": blueprint.get("styleTags", []), "rhetoricPatterns": blueprint.get("rhetoricPatterns", []),
        "sceneRules": blueprint.get("sceneRules", []), "genreGoal": blueprint.get("genreGoal", {}).get("purpose", ""),
        "terms": sorted(set(terms))[:80],
    }


def _extract_source_format_profile(path: Path, template_id: str) -> dict[str, Any]:
    """提取 DOCX 源模板中的实际字体、字号、字重、间距、缩进和对齐观测值。"""
    if path.suffix.lower() != ".docx" or not path.exists():
        return {}
    try:
        doc = Document(str(path))
        structure = build_structure(str(path), template_id)
        role_by_source = {item.get("sourceIndex", item.get("index")): item.get("role") for item in structure}
        observations: dict[str, list[dict[str, Any]]] = {}
        align_names = {None: "left", 0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
        for index, paragraph in enumerate(doc.paragraphs):
            text = paragraph.text.strip()
            if not text:
                continue
            role = role_by_source.get(index, "body")
            run = next((item for item in paragraph.runs if item.text.strip()), None)
            font_east = ""
            if run is not None and run._element.rPr is not None:
                fonts = run._element.rPr.find(qn("w:rFonts"))
                if fonts is not None:
                    font_east = fonts.get(qn("w:eastAsia")) or ""
            fmt = paragraph.paragraph_format
            observations.setdefault(role, []).append({
                "font": font_east or (run.font.name if run else "") or "",
                "sizePt": run.font.size.pt if run and run.font.size else None,
                "bold": bool(run.font.bold) if run else False,
                "align": align_names.get(paragraph.alignment, "left"),
                "firstLinePt": fmt.first_line_indent.pt if fmt.first_line_indent else 0,
                "leftIndentPt": fmt.left_indent.pt if fmt.left_indent else 0,
                "rightIndentPt": fmt.right_indent.pt if fmt.right_indent else 0,
                "lineSpacing": float(fmt.line_spacing.pt) if hasattr(fmt.line_spacing, "pt") else fmt.line_spacing,
                "spaceBeforePt": fmt.space_before.pt if fmt.space_before else 0,
                "spaceAfterPt": fmt.space_after.pt if fmt.space_after else 0,
            })
        section = doc.sections[0] if doc.sections else None
        return {
            "roles": observations,
            "page": ({"widthCm": section.page_width.cm, "heightCm": section.page_height.cm,
                      "topCm": section.top_margin.cm, "bottomCm": section.bottom_margin.cm,
                      "leftCm": section.left_margin.cm, "rightCm": section.right_margin.cm} if section else {}),
        }
    except Exception as exc:
        return {"warning": str(exc)}


def _parse_json_object(content: str) -> dict[str, Any]:
    text = (content or "").strip()
    if "```" in text:
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.IGNORECASE | re.DOTALL).strip()
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise RuntimeError("模型未返回有效 JSON")
    try:
        payload = json.loads(text[start:end + 1])
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"JSON 解析失败：{exc}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("模型返回的 JSON 不是对象")
    return payload


def enrich_structure_blueprint_with_ai(
    blueprint: dict[str, Any],
    reference_text: str,
    request_url: str,
    api_key: str,
    model_name: str,
) -> dict[str, Any]:
    """模板入库时用 AI 补充段落职责、写法、行文范式与场景规则（JSON 为下游主格式）。"""
    sample = reference_text[:5000]
    structure_preview = blueprint.get("structureText") or render_structure_text(blueprint)
    messages = [
        {
            "role": "system",
            "content": (
                "你是公文模板结构分析器。根据范文提炼可复用的框架特征，输出 JSON。"
                "只描述结构职责、可执行写法、格式约束与行文范式；写法须可指导后续用新材料成文。"
            ),
        },
        {
            "role": "user",
            "content": (
                f"已有规则骨架：\n{structure_preview}\n\n"
                f"范文节选（仅供提炼体例，输出抽象特征）：\n{sample}\n\n"
                "返回 JSON："
                '{"organization":{"opening":"","development":"","closing":"","style":[]},'
                '"sections":[{"order":1,"role":"","function":"","writingMethod":"","formatRules":[],'
                '"constraintTags":[],"paragraphFunction":""}],'
                '"styleTags":[],"rhetoricPatterns":[],"sceneRules":[]}\n'
                "sections 的 order/role 必须与已有骨架一致，只 enrich 职责与写法描述。"
                "writingMethod 写成可执行的材料成文指引：用写法动词说明如何将该段职责落实为对材料事实的表达"
                "（如提炼标题、归纳背景、分项列示、因果推进、收束结论），便于下游按新材料独立成文。"
            ),
        },
    ]
    result = call_chat_model_result(request_url, api_key, model_name, messages, temperature=0.0, max_tokens=2048)
    enriched = _parse_json_object(result.get("content") or "")
    for section in blueprint.get("sections") or []:
        ai_section = next(
            (item for item in enriched.get("sections") or [] if item.get("order") == section.get("order")),
            None,
        )
        if not ai_section:
            continue
        for key in ("function", "writingMethod", "formatRules", "constraintTags", "paragraphFunction"):
            if ai_section.get(key):
                section[key] = ai_section[key]
    if enriched.get("organization"):
        blueprint["organization"] = {
            **(blueprint.get("organization") or {}),
            **{k: v for k, v in enriched["organization"].items() if v},
        }
    if enriched.get("styleTags"):
        blueprint["styleTags"] = enriched["styleTags"]
    if enriched.get("rhetoricPatterns"):
        blueprint["rhetoricPatterns"] = enriched["rhetoricPatterns"]
    if enriched.get("sceneRules"):
        blueprint["sceneRules"] = enriched["sceneRules"]
    blueprint["structureText"] = render_structure_text(blueprint)
    blueprint["enrichedByAi"] = True
    return blueprint


def build_rule_based_material_brief(inventory: dict[str, Any]) -> tuple[str, list[dict[str, Any]]]:
    facts: list[dict[str, Any]] = []
    lines = ["【材料事实清单】"]
    for item in inventory.get("items") or []:
        if item.get("sourceType") in {"header", "footer"}:
            continue
        text = str(item.get("text") or "").strip()
        if not text or RE_EMPTY_MATERIAL_VALUE.match(text):
            continue
        fact_id = str(item.get("id") or f"F{len(facts) + 1}")
        facts.append({
            "id": fact_id,
            "category": str(item.get("sourceType") or "body"),
            "content": text,
            "location": str(item.get("location") or ""),
        })
        lines.append(f"[{fact_id} | {item.get('sourceType')}] {text}")
    brief = "\n".join(lines).strip()
    return brief, facts


def process_upload_material_with_ai(
    inventory: dict[str, Any],
    prompt: str,
    template_label: str,
    template_id: str,
    request_url: str,
    api_key: str,
    model_name: str,
    material_zones: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """上传材料预处理：归纳事实、润色表述，形成可供依体例成文的表达底稿。"""
    raw_text = inventory_to_text(inventory)
    zones = material_zones or analyze_material_zones(inventory, None, template_id)
    zones_text = render_material_zones_text(zones)
    fallback_brief, fallback_facts = build_rule_based_material_brief(inventory)
    messages = [
        {
            "role": "system",
            "content": (
                "你是公文材料整理助手。把原始材料整理为可供成文使用的事实清单与润色底稿。"
                "工作方式：归纳要点、重组句式、压缩冗余，形成表达底稿；"
                "letterBrief 是润写后的主件底稿，不是原文逐句照录。"
                "先识别主件与附件区：主件写入 letterBrief，附件写入 attachmentBrief。"
                "title 必须是整篇主标题（如会议名称、公文事由名），"
                "不要把「一、二、三…」章节小标题填入 title；这类小标题留在 letterBrief 内作分段线索。"
                "同时提取 recipient、attachmentHead、signUnits、signDate。"
                "返回 JSON。"
            ),
        },
        {
            "role": "user",
            "content": (
                f"目标文种：{template_label}（{template_id}）\n"
                f"用户要求：{prompt or '形成正式公文'}\n\n"
                f"{zones_text}\n\n"
                f"原始材料：\n{raw_text[:min(len(raw_text), max(12000, min(28000, MAX_DIRECT_INPUT_CHARS)))]}\n\n"
                "返回："
                '{"summary":"100字内摘要",'
                '"title":"整篇主标题，勿用一、二级小标题",'
                '"recipient":"","attachmentHead":"","signUnits":[],"signDate":"",'
                '"facts":[{"id":"F1","zone":"letter|attachment|signature","category":"","content":""}],'
                '"letterBrief":"主件润写底稿：保留关键事实与数据，换用公文表述重组，篇幅可明显短于原文",'
                '"attachmentBrief":"附件区摘要，仅保留附件名称与关键背景"}'
            ),
        },
    ]
    try:
        # 长材料需要更大输出预算，否则整理环节易截断而退回原文清单
        prep_tokens = 8192 if len(raw_text) >= 6000 else 4096
        result = call_chat_model_result(
            request_url, api_key, model_name, messages, temperature=0.2, max_tokens=prep_tokens,
        )
        payload = _parse_json_object(result.get("content") or "")
        brief = str(payload.get("letterBrief") or payload.get("briefForDrafting") or "").strip()
        facts = payload.get("facts") if isinstance(payload.get("facts"), list) else fallback_facts
        summary = str(payload.get("summary") or zones.get("summary") or "").strip()
        if not brief:
            raise RuntimeError("材料整理未返回 letterBrief")
        ai_title = str(payload.get("title") or "").strip()
        if looks_like_section_heading(ai_title):
            ai_title = ""
        zone_title = str(zones.get("titleText") or "").strip()
        if looks_like_section_heading(zone_title):
            zone_title = ""
        return {
            "summary": summary,
            "facts": facts,
            "briefForDrafting": brief,
            "attachmentBrief": str(payload.get("attachmentBrief") or "").strip(),
            "title": ai_title or zone_title,
            "recipient": str(payload.get("recipient") or zones.get("recipientText") or "").strip(),
            "attachmentHead": str(payload.get("attachmentHead") or zones.get("attachmentHeadText") or "").strip(),
            "signUnits": payload.get("signUnits") if isinstance(payload.get("signUnits"), list) else zones.get("signUnits") or [],
            "signDate": str(payload.get("signDate") or zones.get("signDateText") or "").strip(),
            "rawText": raw_text,
            "materialZones": zones,
            "processingMode": "ai_structured_brief",
        }
    except Exception:
        zone_title = str(zones.get("titleText") or "").strip()
        if looks_like_section_heading(zone_title):
            zone_title = ""
        return {
            "summary": zones.get("summary") or "",
            "facts": fallback_facts,
            "briefForDrafting": fallback_brief,
            "attachmentBrief": "",
            "title": zone_title,
            "recipient": zones.get("recipientText") or "",
            "attachmentHead": zones.get("attachmentHeadText") or "",
            "signUnits": zones.get("signUnits") or [],
            "signDate": zones.get("signDateText") or "",
            "rawText": raw_text,
            "materialZones": zones,
            "processingMode": "rule_based_brief",
        }


def reference_structure_payload(reference: dict[str, Any]) -> dict[str, Any]:
    """仅返回可调用的结构/范式字段，剥离范文原文与禁用字段。"""
    positive_anti_plagiarism = {
        "reuseStructure": True,
        "reuseRhetoric": True,
        "reuseVerbatim": False,
        "rule": "以结构化段落职责与写法为体例坐标，用材料事实完成各段独立成文表达",
    }
    blueprint = reference.get("structureJson") or {}
    template_id = str(
        blueprint.get("templateId")
        or reference.get("templateId")
        or ""
    )
    genre_profile = build_genre_writing_profile(template_id) if template_id else {}
    callable_keys = blueprint.get("callableFields") or [
        "sections", "headingSkeleton", "roleSequence", "requiredRoles",
        "organization", "styleTags", "rhetoricPatterns", "sceneRules",
        "structureText", "antiPlagiarism", "genreGoal", "formatProfile", "featureVector",
    ]
    if blueprint:
        payload = {key: blueprint[key] for key in callable_keys if key in blueprint}
        # 段落内去掉 patternHint 原文痕迹，只保留职责与约束
        sections = []
        for section in payload.get("sections") or []:
            sections.append({
                "order": section.get("order"),
                "role": section.get("role"),
                "label": section.get("label"),
                "function": section.get("function") or section.get("paragraphFunction"),
                "writingMethod": section.get("writingMethod"),
                "formatRules": section.get("formatRules") or [],
                "constraintTags": section.get("constraintTags") or [],
                "paragraphFunction": section.get("paragraphFunction") or section.get("function"),
            })
        if sections:
            payload["sections"] = sections
        # 运行时统一正向体例表述，兼容库内旧版“仿写/复刻”文案
        payload["antiPlagiarism"] = {
            **positive_anti_plagiarism,
            **({k: v for k, v in (payload.get("antiPlagiarism") or {}).items()
                if k in {"reuseStructure", "reuseRhetoric", "reuseVerbatim"} and v is not None}),
            "rule": positive_anti_plagiarism["rule"],
        }
        # 运行时注入文种写作品格，避免仅依赖范文 roleSequence 误推层级标题
        if genre_profile:
            payload["writingProfile"] = {
                "mode": genre_profile.get("mode"),
                "preferredRoles": genre_profile.get("preferredRoles"),
                "purpose": genre_profile.get("purpose"),
                "organization": genre_profile.get("organization"),
                "headingHint": genre_profile.get("headingHint"),
            }
            if genre_profile.get("mode") == "layout_sections":
                payload["headingSkeleton"] = []
                payload["organization"] = {
                    **(payload.get("organization") or {}),
                    **(genre_profile.get("organization") or {}),
                }
        payload.setdefault("requiredRoles", ["title"])
        return payload
    roles = reference.get("roles") or []
    payload = {
        "roleSequence": roles,
        "requiredRoles": ["title"],
        "headingSkeleton": [role for role in roles if role in {"h1", "h2", "h3", "h4"}],
        "structureText": " -> ".join(ROLE_NAMES.get(role, role) for role in roles),
        "antiPlagiarism": dict(positive_anti_plagiarism),
    }
    if genre_profile:
        payload["writingProfile"] = {
            "mode": genre_profile.get("mode"),
            "preferredRoles": genre_profile.get("preferredRoles"),
            "purpose": genre_profile.get("purpose"),
            "organization": genre_profile.get("organization"),
            "headingHint": genre_profile.get("headingHint"),
        }
        if genre_profile.get("mode") == "layout_sections":
            payload["headingSkeleton"] = []
    return payload


def _reference_role_sequence(path: Path, template_id: str, text: str) -> list[str]:
    if path.suffix.lower() == ".docx":
        try:
            items = build_structure(str(path), template_id)
            roles = [item["role"] for item in items]
            if roles:
                return _compact_roles(roles)
        except Exception:
            pass

    roles: list[str] = []
    for line in text.splitlines():
        value = line.strip()
        if not value:
            continue
        if not roles:
            roles.append("title")
        elif re.match(r"^[一二三四五六七八九十]+[、.]", value):
            roles.append("h1")
        elif re.match(r"^[（(][一二三四五六七八九十]+[）)]", value):
            roles.append("h2")
        elif re.match(r"^\d+[.．、]", value):
            roles.append("h3")
        else:
            roles.append("body")
    return _compact_roles(roles)


def _reference_text(inventory: dict[str, Any]) -> str:
    return "\n\n".join(
        str(item.get("text") or "").strip()
        for item in inventory.get("items", [])
        if str(item.get("text") or "").strip()
    )


def _ingest_reference(
    path: Path,
    template_id: str,
    *,
    enrich_with_ai: bool = False,
    model_config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    base = {
        "templateId": template_id,
        "sourceDir": str(path.parent),
        "filename": path.name,
        "filePath": str(path),
        "suffix": path.suffix.lower(),
        "status": "unreadable",
        "warning": "",
        "text": "",
        "textChars": 0,
        "roles": [],
    }
    if path.suffix.lower() not in READABLE_REFERENCE_SUFFIXES:
        if path.suffix.lower() in UNREADABLE_REFERENCE_SUFFIXES:
            base["warning"] = f"{path.suffix} 暂无可靠解析器，请先转换为 DOCX 或 PDF 后再同步"
        else:
            base["warning"] = f"{path.suffix} 暂无可靠解析器，请先转换为 DOCX 或 PDF"
        return base
    try:
        inventory = extract_document_inventory(path)
        text = _reference_text(inventory).strip()
        quality = assess_material_content(inventory)
        warnings = list(inventory.get("warnings") or [])
        if quality["status"] != "ready" or not text:
            warnings.append(quality["reason"] or "未提取到范文正文")
            base["warning"] = "；".join(warnings)
            return base
        base.update({
            "status": "ready",
            "warning": "；".join(warnings),
            "text": text,
            "textChars": len(text),
            "roles": _reference_role_sequence(path, template_id, text),
        })
        blueprint = build_rule_based_structure_blueprint(
            template_id, path.name, base["roles"], text, path,
        )
        if enrich_with_ai and model_config and model_config.get("apiKeyConfigured"):
            try:
                blueprint = enrich_structure_blueprint_with_ai(
                    blueprint,
                    text,
                    str(model_config.get("requestUrl") or ""),
                    str(model_config.get("apiKey") or os.getenv("DEEPSEEK_API_KEY", "")),
                    str(model_config.get("modelName") or os.getenv("DEEPSEEK_MODEL", "")),
                )
            except Exception as exc:
                blueprint["enrichWarning"] = str(exc)
        base["structureJson"] = blueprint
    except Exception as exc:
        base["warning"] = f"范文解析失败：{exc}"
    return base


def _build_corpus_profile(references: list[dict[str, Any]]) -> dict[str, Any]:
    ready = [item for item in references if item["status"] == "ready"]
    sequences = [tuple(item["roles"]) for item in ready if item.get("roles")]
    skipped_details = [
        {"filename": item["filename"], "reason": item["warning"] or "未提取到范文正文"}
        for item in references if item["status"] != "ready"
    ]
    if not ready:
        return {
            "schemaVersion": 3,
            "status": "unavailable",
            "analyzedFileCount": 0,
            "availableReferenceCount": 0,
            "dominantSequence": [],
            "headingSkeleton": [],
            "variants": [],
            "skippedFiles": [item["filename"] for item in skipped_details],
            "skippedDetails": skipped_details,
            "summary": "没有可解析范文，生成时将被阻止；请先将范文转换为 DOCX 或 PDF。",
        }

    counts = Counter(sequences)
    ranked = sorted(counts.items(), key=lambda item: (-item[1], -len(item[0]), item[0]))
    dominant = ranked[0][0] if ranked else ()
    dominant_count = ranked[0][1] if ranked else 0
    variants = [
        {"roles": list(sequence), "count": count}
        for sequence, count in ranked[1:4]
    ]
    role_presence = Counter(role for sequence in sequences for role in set(sequence))
    common_roles = [
        role for role, count in role_presence.items()
        if count * 2 >= max(len(sequences), 1)
    ]
    dominant_names = " -> ".join(ROLE_NAMES.get(role, role) for role in dominant) or "由具体范文确定"
    heading_skeleton = [role for role in dominant if role in {"h1", "h2", "h3", "h4"}]
    return {
            "schemaVersion": 3,
            "status": "ready",
        "analyzedFileCount": len(ready),
        "availableReferenceCount": len(ready),
        "structuredReferenceCount": len(sequences),
        "dominantSequence": list(dominant),
        "headingSkeleton": heading_skeleton,
        "dominantCount": dominant_count,
        "variants": variants,
        "commonRoles": common_roles,
        "skippedFiles": [item["filename"] for item in skipped_details],
        "skippedDetails": skipped_details,
        "summary": f"已完整读取 {len(ready)} 份范文；公共结构为：{dominant_names}。生成时自动选择最相近的一份。",
    }


def sync_template_library(enrich_structures: bool = False) -> list[dict[str, Any]]:
    init_template_db()
    catalog_by_label = {item["label"]: item for item in get_template_catalog()}
    templates: list[dict[str, Any]] = []
    all_references: list[dict[str, Any]] = []
    model_config = get_model_config() if enrich_structures else None
    if not TEMPLATE_SOURCE_DIR.exists():
        return load_templates_from_db()

    for folder in sorted([p for p in TEMPLATE_SOURCE_DIR.iterdir() if p.is_dir()], key=lambda p: p.name):
        files = sorted([
            item for item in folder.iterdir()
            if item.is_file()
            and not item.name.startswith("~$")
            and item.suffix.lower() in REFERENCE_SUFFIXES
        ], key=lambda item: item.name)
        template_id = guess_template_id("", folder.name)
        catalog_item = next((v for v in catalog_by_label.values() if v["id"] == template_id), None)
        references = [
            _ingest_reference(
                path, template_id,
                enrich_with_ai=enrich_structures,
                model_config=model_config,
            )
            for path in files
        ]
        ready = [item for item in references if item["status"] == "ready"]
        corpus_profile = _build_corpus_profile(references)
        sample = ready[0] if ready else None
        all_references.extend(references)
        templates.append({
            "id": template_id,
            "name": folder.name,
            "label": catalog_item["label"] if catalog_item else folder.name,
            "fileCount": len(files),
            "sourceDir": str(folder),
            "sampleFile": sample["filename"] if sample else "",
            "samplePath": sample["filePath"] if sample else "",
            "sampleText": sample["text"] if sample else "",
            "files": [item.name for item in files],
            "corpusProfile": corpus_profile,
        })

    from datetime import datetime
    now = datetime.now().isoformat(timespec="seconds")
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.execute("DELETE FROM template_references")
        conn.execute("DELETE FROM document_templates")
        conn.executemany(
            """
            INSERT INTO document_templates (
                id, name, label, file_count, source_dir, sample_file,
                sample_path, sample_text, files_json, corpus_profile_json, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id, source_dir) DO UPDATE SET
                name=excluded.name,
                label=excluded.label,
                file_count=excluded.file_count,
                sample_file=excluded.sample_file,
                sample_path=excluded.sample_path,
                sample_text=excluded.sample_text,
                files_json=excluded.files_json,
                corpus_profile_json=excluded.corpus_profile_json,
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
                    json.dumps(item["corpusProfile"], ensure_ascii=False),
                    now,
                )
                for item in templates
            ],
        )
        conn.executemany(
            """
            INSERT INTO template_references (
                template_id, source_dir, filename, file_path, suffix, status,
                warning, text, text_chars, roles_json, structure_json, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    item["templateId"], item["sourceDir"], item["filename"],
                    item["filePath"], item["suffix"], item["status"], item["warning"],
                    item["text"], item["textChars"],
                    json.dumps(item["roles"], ensure_ascii=False),
                    json.dumps(item.get("structureJson") or {}, ensure_ascii=False),
                    now,
                )
                for item in all_references
            ],
        )
    return sorted(load_templates_from_db(), key=_template_sort_key)


def scan_template_library() -> list[dict[str, Any]]:
    templates = load_templates_from_db()
    if templates:
        with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
            reference_count = conn.execute("SELECT COUNT(*) FROM template_references").fetchone()[0]
        needs_upgrade = any(
            int((item.get("corpusProfile") or {}).get("schemaVersion", 0) or 0) < 3
            for item in templates
        )
        if reference_count and not needs_upgrade:
            return templates
    return sync_template_library()


def get_template_by_id(template_id: str | None) -> dict[str, Any] | None:
    if not template_id:
        return None
    return next((item for item in scan_template_library() if item["id"] == template_id), None)


def public_template_view(template: dict[str, Any]) -> dict[str, Any]:
    """模板正文仅供后端生成使用，模板接口只暴露索引元数据。"""
    hidden = {"sourceDir", "samplePath", "sampleText"}
    return {key: value for key, value in template.items() if key not in hidden}


def load_template_references(template_id: str, source_dir: str) -> list[dict[str, Any]]:
    init_template_db()
    with sqlite3.connect(TEMPLATE_DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT filename, file_path, suffix, status, warning, text, text_chars, roles_json, structure_json
            FROM template_references
            WHERE template_id = ? AND source_dir = ?
            ORDER BY filename
            """,
            (template_id, source_dir),
        ).fetchall()
    references = []
    for row in rows:
        try:
            roles = json.loads(row["roles_json"] or "[]")
        except json.JSONDecodeError:
            roles = []
        try:
            structure_json = json.loads(row["structure_json"] or "{}")
        except json.JSONDecodeError:
            structure_json = {}
        if (not structure_json or int(structure_json.get("version", 0) or 0) < 3) and roles:
            structure_json = build_rule_based_structure_blueprint(
                template_id, row["filename"], roles, row["text"] or "",
                Path(row["file_path"]) if row["file_path"] else None,
            )
        references.append({
            "templateId": template_id,
            "filename": row["filename"],
            "filePath": row["file_path"],
            "suffix": row["suffix"],
            "status": row["status"],
            "warning": row["warning"],
            "text": row["text"],
            "textChars": row["text_chars"],
            "roles": roles,
            "structureJson": structure_json,
        })
    return references


def _text_ngrams(text: str, size: int = 2) -> Counter[str]:
    normalized = "".join(re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", text.lower()))
    if len(normalized) < size:
        return Counter({normalized: 1}) if normalized else Counter()
    return Counter(normalized[index:index + size] for index in range(len(normalized) - size + 1))


def _reference_similarity_scores(query: str, references: list[dict[str, Any]]) -> list[float]:
    query_counts = _text_ngrams(query)
    document_counts = [
        _text_ngrams(f"{item['filename']}\n{item['filename']}\n{item['text']}")
        for item in references
    ]
    document_frequency = Counter(
        gram for counts in document_counts for gram in counts
    )
    total_documents = len(document_counts)

    def vector(counts: Counter[str]) -> dict[str, float]:
        return {
            gram: count * (math.log((total_documents + 1) / (document_frequency.get(gram, 0) + 1)) + 1)
            for gram, count in counts.items()
        }

    query_vector = vector(query_counts)
    query_norm = math.sqrt(sum(value * value for value in query_vector.values()))
    query_weight = sum(query_vector.values())
    scores = []
    for counts in document_counts:
        candidate_vector = vector(counts)
        candidate_norm = math.sqrt(sum(value * value for value in candidate_vector.values()))
        dot = sum(value * candidate_vector.get(gram, 0.0) for gram, value in query_vector.items())
        cosine = dot / (query_norm * candidate_norm) if query_norm and candidate_norm else 0.0
        coverage = (
            sum(value for gram, value in query_vector.items() if gram in candidate_vector) / query_weight
            if query_weight else 0.0
        )
        scores.append(coverage * 0.8 + cosine * 0.2)
    return scores


def _template_feature_score(query: str, prompt: str, reference: dict[str, Any], template_id: str) -> tuple[float, dict[str, float]]:
    feature = (reference.get("structureJson") or {}).get("featureVector") or {}
    query_lines = [line.strip() for line in query.splitlines() if line.strip()]
    query_roles = _reference_role_sequence(Path("__missing__"), template_id, query) if query else []
    query_terms = set(re.findall(r"[\u4e00-\u9fff]{2,8}", query))
    candidate_terms = set(feature.get("terms") or [])
    text_score = min(len(query) / max(float(feature.get("textChars") or 1), 1), 1.0)
    paragraph_score = 1.0 - min(abs(len(query_lines) - int(feature.get("paragraphCount") or 1)) / max(len(query_lines), 1), 1.0)
    role_score = 1.0 if not query_roles or not feature.get("roleSequence") else SequenceMatcher(None, query_roles, feature["roleSequence"]).ratio()
    term_score = len(query_terms & candidate_terms) / max(len(query_terms), 1)
    genre_score = 1.0 if any(keyword in f"{prompt}{query}" for keyword in _genre_keywords(template_id)) else 0.5
    scores = {"textLength": text_score, "paragraphDistribution": paragraph_score, "roleStructure": role_score, "terminology": term_score, "genre": genre_score}
    return round(sum(scores.values()) / len(scores), 4), scores


def _genre_keywords(template_id: str) -> list[str]:
    return {
        "report": ["汇报", "情况", "总结", "进展"], "request": ["请示", "请求", "批准"],
        "notice": ["通知", "部署", "要求", "传达"], "letter": ["函", "商请", "协助"],
        "opinion": ["意见", "实施", "建议"], "minutes": ["会议", "纪要", "议定"],
    }.get(template_id, [])


def select_reference_for_material(
    template: dict[str, Any], prompt: str, materials: list[dict[str, Any]],
) -> dict[str, Any]:
    references = [
        item for item in load_template_references(template["id"], template["sourceDir"])
        if item["status"] == "ready" and item["text"].strip()
    ]
    if not references:
        raise RuntimeError(
            f"模板“{template.get('name') or template.get('label')}”没有可解析范文，"
            "已按设置停止生成。请先将该文种范文转换为 DOCX 或 PDF 并重新同步模板库"
        )
    material_bits = []
    for item in materials:
        text = str(item.get("rawText") or item.get("text") or "").strip()
        if text:
            # 用材料原文匹配体例样本，避免 brief/模板痕迹干扰；截断以控制打分成本
            material_bits.append(text[:8000])
    query = "\n".join([prompt, *material_bits]).strip() or prompt
    base_scores = _reference_similarity_scores(query, references)
    scored = []
    for index, reference in enumerate(references):
        multi_score, dimensions = _template_feature_score(query, prompt, reference, template["id"])
        final_score = round(base_scores[index] * 0.35 + multi_score * 0.65, 4)
        scored.append((final_score, dimensions))
    selected_index = max(range(len(references)), key=lambda index: (scored[index][0], -index))
    selected = dict(references[selected_index])
    selected["similarity"] = scored[selected_index][0]
    selected["matchDimensions"] = scored[selected_index][1]
    selected["templateId"] = template["id"]
    if len(selected["text"]) > MAX_REFERENCE_INPUT_CHARS:
        raise RuntimeError(
            f"自动匹配范文“{selected['filename']}”共 {len(selected['text'])} 字，超过范文输入上限 "
            f"{MAX_REFERENCE_INPUT_CHARS} 字；请精简范文或提高 AI_MAX_REFERENCE_INPUT_CHARS"
        )
    return selected


def _reference_paragraphs(reference: dict[str, Any]) -> list[dict[str, str]]:
    paragraphs = []
    for index, text in enumerate(
        part.strip() for part in str(reference.get("text") or "").split("\n\n")
    ):
        if text:
            paragraphs.append({"id": f"R{index + 1}", "text": text})
    return paragraphs


def _source_ids(materials: list[dict[str, Any]]) -> list[str]:
    ids: list[str] = []
    for item in materials:
        for fact in item.get("structuredFacts") or []:
            fact_id = str(fact.get("id") or "").strip()
            if fact_id and fact_id not in ids:
                ids.append(fact_id)
        for source_id in _source_ids_in_text(item.get("text", "")):
            if source_id not in ids:
                ids.append(source_id)
    if not ids:
        ids.append("prompt")
    return ids


def _heading_design_hint(structure_payload: dict[str, Any], template_id: str = "") -> str:
    """按文种写作品格给出标题/版式设计提示。"""
    profile = build_genre_writing_profile(template_id) if template_id else {}
    if profile.get("headingHint"):
        return str(profile["headingHint"])
    skeleton = structure_payload.get("headingSkeleton") or []
    h1_count = sum(1 for role in skeleton if role == "h1")
    h2_count = sum(1 for role in skeleton if role == "h2")
    if not skeleton and not h1_count:
        return "按材料复杂度决定是否分块：内容少时可连续正文，内容多时用一级「一、二、三」划分大段。"
    parts = [f"体例样本主体约 {max(h1_count, 1)} 个一级大段"]
    if h2_count:
        parts.append("大段内可用「（一）（二）」作二级细分")
    parts.append("标题数量与措辞以材料事实为准，可增可减，不必与样本一一对应")
    return "；".join(parts) + "。"


def _letter_source_ids(materials: list[dict[str, Any]]) -> list[str]:
    """主件正文区事实编号，供结构计划与成文取材。"""
    ids: list[str] = []
    for item in materials:
        zones = item.get("materialZones") or {}
        for fact_id in zones.get("letterFactIds") or []:
            if fact_id not in ids:
                ids.append(fact_id)
    return ids or _source_ids(materials)


def _material_context_text(materials: list[dict[str, Any]]) -> str:
    parts = []
    for index, item in enumerate(materials, start=1):
        zones = item.get("materialZones") or {}
        zones_text = render_material_zones_text(zones) if zones else ""
        attachment_note = ""
        if item.get("attachmentBrief"):
            attachment_note = f"\n附件区摘要（不写入函件正文）：\n{item['attachmentBrief'][:1500]}"
        mode = item.get("processingMode") or ""
        rewrite_guide = (
            "成文取用说明：以下为主件事实来源。请据此润写重组正式公文："
            "保留关键事实与数据，换用公文句式与概括层次；"
            "输出完整篇章（独立主标题、正文、材料具备时的落款），避免原文逐句照录。"
        )
        if mode == "rule_based_brief":
            rewrite_guide += "当前为规则事实清单，更需归纳重组后再成文。"
        elif mode == "ai_structured_brief":
            rewrite_guide += "当前为已润色底稿，请继续按文种体例写成完整公文。"
        parts.append(
            "\n".join([
                f"[Material {index}: {item['name']}]",
                rewrite_guide,
                zones_text,
                f"主件表达底稿：\n{item.get('text', '')}",
                attachment_note,
            ]).strip()
        )
    return "\n\n".join(parts)


def build_transformation_plan_messages(
    prompt: str, materials: list[dict[str, Any]], template: dict[str, Any],
    reference: dict[str, Any],
) -> list[dict[str, str]]:
    structure_payload = reference_structure_payload(reference)
    reference_outline = json.dumps(structure_payload, ensure_ascii=False, indent=2)
    material_text = _material_context_text(materials)
    letter_ids = _letter_source_ids(materials)
    allowed_ids = ", ".join(letter_ids)
    zones_notes = "\n\n".join(
        render_material_zones_text(item.get("materialZones") or {})
        for item in materials
        if item.get("materialZones")
    )
    template_id = str(template.get("id") or "")
    genre_profile = build_genre_writing_profile(template_id)
    preferred = "、".join(genre_profile.get("preferredRoles") or [])
    system = (
        "你是公文结构策划助手。先识别目标文种的写作品格，再为上传材料设计可执行的成文段落蓝图（JSON）。"
        f"当前文种「{genre_profile.get('label') or template_id}」：{genre_profile.get('purpose', '')}。"
        f"{genre_profile.get('planGuide', '')}"
        "体例 JSON 提供节奏与段落职责；标题文字与分段数量由材料内容决定。"
        "蓝图完整性：必须包含独立 title；「一、二、三…」规划为 h1；材料有落款则规划 sign_unit/sign_date。"
        "成文对象是主件：附件区材料用于 attachment_head 一行说明。"
        "body 及主体段落的 sourceIds 仅取主件正文区编号；材料未覆盖的段落可省略。"
        "返回 JSON 对象，不写 Markdown。"
    )
    user = (
        f"目标文种：{genre_profile.get('label') or template.get('label') or template.get('name')}（{template_id}）\n"
        f"文种用途：{genre_profile.get('purpose', '')}\n"
        f"写作品格：{genre_profile.get('mode', 'adaptive')}；常用角色：{preferred}\n"
        f"组织方式：{json.dumps(genre_profile.get('organization') or {}, ensure_ascii=False)}\n"
        f"用户要求：{prompt or '根据上传材料形成正式公文'}\n"
        f"主件正文区可用来源编号：{allowed_ids}\n\n"
        f"{zones_notes}\n\n"
        "模板结构化框架（JSON，文种体例坐标）：\n<template_structure>\n"
        f"{reference_outline}\n</template_structure>\n\n"
        "上传材料（主件与附件已分区）：\n<source>\n"
        f"{material_text or prompt}\n</source>\n\n"
        "返回以下结构：\n"
        '{"documentPurpose":"文种用途",'
        '"organization":{"opening":"开头组织方式", "development":"主体展开方式", "closing":"结尾方式", "style":["句式或衔接特点"]},'
        '"targetSections":[{"order":1,"role":"title|subtitle|recipient|body|h1|h2|h3|h4|attachment_head|attachment_other|sign_unit|sign_date|sign_contact|other",'
        '"function":"本段职责", "writingMethod":"本段材料成文指引", "sourceIds":["来源编号"]}]}\n'
        f"优先按常用角色设计：{preferred}。"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def parse_transformation_plan(content: str, materials: list[dict[str, Any]]) -> dict[str, Any]:
    text = (content or "").strip()
    if "```" in text:
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.IGNORECASE | re.DOTALL).strip()
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise RuntimeError("范文结构分析未返回有效JSON")
    try:
        plan = json.loads(text[start:end + 1])
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"范文结构分析JSON解析失败：{exc}") from exc
    sections = plan.get("targetSections")
    organization = plan.get("organization")
    if not isinstance(sections, list) or not sections or not isinstance(organization, dict):
        raise RuntimeError("范文结构分析缺少组织方式或目标段落映射")
    allowed_roles = ALLOWED_DRAFT_ROLES
    valid_ids = set(_source_ids(materials))
    normalized_sections = []
    for index, section in enumerate(sections, start=1):
        if not isinstance(section, dict) or section.get("role") not in allowed_roles:
            continue
        source_ids = [str(item) for item in (section.get("sourceIds") or []) if str(item) in valid_ids]
        normalized_sections.append({
            "order": int(section.get("order") or index),
            "role": section["role"],
            "function": str(section.get("function") or "完成本段公文职责，组织对应材料事实"),
            "writingMethod": str(section.get("writingMethod") or "按本段写法，用对应 sourceIds 事实完成表达"),
            "sourceIds": source_ids,
        })
    if not normalized_sections:
        raise RuntimeError("范文结构分析没有可用目标段落")
    plan["targetSections"] = normalized_sections
    plan["organization"] = {
        "opening": str(organization.get("opening") or "按文种开头职责组织材料事实"),
        "development": str(organization.get("development") or "按文种主体职责展开材料事实"),
        "closing": str(organization.get("closing") or "按文种结尾职责收束材料事实"),
        "style": [str(item) for item in (organization.get("style") or []) if str(item)],
    }
    plan["documentPurpose"] = str(plan.get("documentPurpose") or "按目标文种形成正式公文")
    return plan


def build_ai_messages(
    prompt: str, materials: list[dict[str, Any]], template: dict | None,
    reference: dict[str, Any], transformation_plan: dict[str, Any],
    *,
    omit_reference_body: bool = True,
    rewrite_mode: bool = False,
    rewrite_reason: str = "",
) -> list[dict[str, str]]:
    structure_payload = reference_structure_payload(reference)
    output_limit = draft_char_limit(materials, reference)
    plan_text = json.dumps(transformation_plan, ensure_ascii=False, indent=2)
    template_id = template.get("id", "") if template else ""
    genre_profile = build_genre_writing_profile(template_id)
    heading_hint = _heading_design_hint(structure_payload, template_id)
    template_note = (
        f"文种：{genre_profile.get('label') or (template or {}).get('label') or '通用公文'}（{template_id or 'generic'}）\n"
        f"文种用途：{genre_profile.get('purpose', '')}\n"
        f"写作品格：{genre_profile.get('mode', 'adaptive')}；常用角色：{'、'.join(genre_profile.get('preferredRoles') or [])}\n"
        f"体例参照样本：{reference['filename']}\n"
        f"匹配分数：{reference.get('similarity', 0):.4f}\n"
        f"版式设计：{heading_hint}"
    )

    material_text = _material_context_text(materials)
    zones_notes = "\n\n".join(
        render_material_zones_text(item.get("materialZones") or {})
        for item in materials
        if item.get("materialZones")
    )
    source_elements = [
        element
        for item in materials
        for element in item.get("sourceElements", [])
    ]
    source_elements_text = "\n".join(
        f"[[{element['role']}]]{element['text']}"
        for element in source_elements
    ) or "无"
    rewrite_note = ""
    if rewrite_mode:
        rewrite_note = (
            "【再生成】保留全部材料事实与结构映射计划，重新组织表达。"
            f"触发原因：{rewrite_reason or '表述与体例样本或材料底稿相近'}。"
            "各段换用不同句式与概括方式；继续遵循当前文种写作品格与版式节奏。"
        )
    chapter_guide = genre_profile.get("chapterGuide") or (
        "按材料主题决定连续正文或分块标题；标题后紧跟正文。"
    )
    example = genre_profile.get("example") or (
        "[[title]]关于×××的公文\n[[body]]……正文……\n"
    )
    system = (
        "你是中文公文撰稿助手。依据上传材料中的事实，写成一篇符合目标文种写作品格、"
        "可直接使用的正式公文。\n\n"
        "【成文原则】\n"
        f"当前文种「{genre_profile.get('label') or template_id}」：{genre_profile.get('purpose', '')}。"
        "事实与数据来自用户要求与上传材料；体例 JSON 提供篇章节奏与格式参照。"
        "成文是对材料的润写重组：换用公文句式与概括方式表达事实，"
        "形成结构完整的正式文稿，而不是把材料原文逐段粘贴。\n"
        "先识别材料分区：主件写入成文主体，附件区长文仅在 [[attachment_head]] 点名。\n\n"
        "【篇章完整性】\n"
        "完整公文按文种输出：独立主标题 [[title]] →（材料有则）[[subtitle]]/[[recipient]] → "
        "主体正文（[[body]] 与按需 [[h1]]/[[h2]]）→（材料有则）[[attachment_head]] → "
        "落款 [[sign_unit]]/[[sign_date]]/[[sign_contact]]。\n"
        "主标题概括整篇主题（如会议名称、事由），居中独立成段；"
        "「一、二、三…」只作 [[h1]]，「（一）（二）」只作 [[h2]]，"
        "二者都不充当主标题。\n\n"
        "【篇章设计】\n"
        f"{chapter_guide}\n"
        "附件说明单独 [[attachment_head]] 一行；落款单位每行一个 [[sign_unit]]，"
        "成文日期独立 [[sign_date]]，联系人及电话可用 [[sign_contact]]。\n\n"
        "【事实边界】\n"
        "输出中的动作、要求、程序、责任与结果均能在主件材料中找到依据；"
        "材料已有的主送、附件说明、落款、日期按原文字保留；"
        "主标题优先采用材料中的整篇标题；若材料仅有章节小标题，则从会议/主题名称提炼 [[title]]，"
        "章节小标题保留为 [[h1]]。\n\n"
        "【输出格式】\n"
        "按结构映射计划与文种常用角色安排段落顺序；每段独立一行，必须以 [[role]] 开头。"
        "role 为 title、subtitle、recipient、body、h1、h2、h3、h4、"
        "attachment_head、attachment_other、sign_unit、sign_date、sign_contact、other。\n"
        "示例：\n"
        f"{example}"
        f"成文上限 {output_limit} 字。只输出上述结构化段落，不写说明或 Markdown。"
        + (f"\n\n{rewrite_note}" if rewrite_note else "")
    )
    structure_json = json.dumps(structure_payload, ensure_ascii=False, indent=2)
    if omit_reference_body or rewrite_mode:
        reference_body = (
            "（体例坐标：参照段落职责、层次深度与写法，用材料事实独立成文。）\n"
            f"{structure_json}"
        )
    else:
        reference_body = (
            f"{structure_json}\n\n---\n体例样本原文（仅供对照篇章骨架）：\n"
            f"{reference.get('text', '')[:3000]}"
        )
    user = (
        f"用户要求：\n{prompt or '根据材料形成正式公文。'}\n\n"
        f"文种标识：{template_id or 'generic'}\n"
        f"自动匹配结果：\n{template_note}\n\n"
        + (f"{zones_notes}\n\n" if zones_notes else "")
        + f"结构映射计划（成文蓝图）：\n<plan>\n{plan_text}\n</plan>\n\n"
        f"模板结构化框架（JSON，文种体例坐标）：\n"
        f"<template_structure>\n{reference_body}\n</template_structure>\n\n"
        f"上传材料中必须保留的结构要素：\n{source_elements_text}\n\n"
        f"成文内容来源（主件与附件已分区）：\n<source>\n"
        f"{material_text or '无上传材料，仅可使用用户要求中明确写出的信息。'}\n</source>"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def is_similarity_failure(error: Exception) -> bool:
    text = str(error)
    return any(marker in text for marker in SIMILARITY_FAILURE_MARKERS)


def build_parse_retry_message(error: Exception) -> str:
    return (
        f"上一次输出无法解析：{error}。请按结构映射计划与 [[role]]正文 格式"
        "重新输出完整公文；每个标题后紧跟正文段，只输出结构化段落。"
    )


def build_similarity_rewrite_messages(
    prompt: str,
    materials: list[dict[str, Any]],
    template: dict[str, Any],
    reference: dict[str, Any],
    transformation_plan: dict[str, Any],
    error: Exception,
    reference_audit: dict[str, Any] | None = None,
    validation_metrics: dict[str, Any] | None = None,
    previous_draft: str = "",
) -> list[dict[str, str]]:
    """相似度风控触发后，按范文结构对材料做语义差异化重构。"""
    audit_warnings = list((reference_audit or {}).get("warnings") or [])
    source_coverage = (validation_metrics or {}).get("sourceCopyCoverage")
    copied_ratio = (validation_metrics or {}).get("copiedParagraphRatio")
    detail_parts = [str(error)]
    if source_coverage is not None:
        detail_parts.append(f"材料重合度={source_coverage}")
    if copied_ratio is not None:
        detail_parts.append(f"段落复制比={copied_ratio}")
    if audit_warnings:
        detail_parts.append("需规避表述：" + "；".join(audit_warnings[:3]))
    rewrite_reason = "；".join(detail_parts)
    messages = build_ai_messages(
        prompt, materials, template, reference, transformation_plan,
        omit_reference_body=True,
        rewrite_mode=True,
        rewrite_reason=rewrite_reason,
    )
    if previous_draft.strip():
        messages.extend([
            {"role": "assistant", "content": previous_draft},
            {
                "role": "user",
                "content": (
                    "请按结构映射计划重新成文：保留全部事实，换用不同句式与概括方式；"
                    "每个标题后写对应正文，h2 在每个 h1 下从（一）重新编号。"
                    "按 [[role]]正文 输出完整公文。"
                ),
            },
        ])
    return messages


def resolve_chat_url(request_url: str) -> str:
    url = (request_url or "").strip().rstrip("/")
    if not url:
        raise ValueError("请填写模型请求地址")
    if url.endswith("/chat/completions"):
        return url
    if url.endswith("/v1"):
        return f"{url}/chat/completions"
    return f"{url}/v1/chat/completions"


def call_chat_model_result(
    request_url: str,
    api_key: str,
    model_name: str,
    messages: list[dict[str, str]],
    temperature: float = 0.2,
    max_tokens: int | None = None,
) -> dict[str, Any]:
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
        "thinking": {"type": "disabled"},
    }
    if max_tokens:
        payload["max_tokens"] = max_tokens
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(resolve_chat_url(request_url), data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=MODEL_HTTP_TIMEOUT_SECONDS) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except TimeoutError as exc:
        raise RuntimeError(
            f"模型接口超时（{MODEL_HTTP_TIMEOUT_SECONDS} 秒）。"
            "高并发时推理排队较长，可提高环境变量 AI_MODEL_HTTP_TIMEOUT_SECONDS 后重试"
        ) from exc
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"模型接口返回 {exc.code}: {detail[:500]}") from exc
    except Exception as exc:
        detail = str(exc)
        if "timed out" in detail.lower() or "timeout" in detail.lower():
            raise RuntimeError(
                f"模型接口超时（{MODEL_HTTP_TIMEOUT_SECONDS} 秒）。"
                "高并发时推理排队较长，可提高环境变量 AI_MODEL_HTTP_TIMEOUT_SECONDS 后重试"
            ) from exc
        raise RuntimeError(f"模型接口请求失败: {exc}") from exc
    try:
        choice = body["choices"][0]
        message = choice.get("message") or {}
        content = message.get("content") or ""
        if isinstance(content, list):
            content = "".join(
                str(part.get("text") or "") if isinstance(part, dict) else str(part)
                for part in content
            )
        return {
            "content": str(content).strip(),
            "finishReason": choice.get("finish_reason"),
            "model": body.get("model", model_name),
            "usage": body.get("usage") or {},
            "reasoningChars": len(str(message.get("reasoning_content") or "")),
        }
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("模型接口响应格式不符合 OpenAI Chat Completions 规范") from exc


def sanitize_draft_text(text: str) -> tuple[str, list[str]]:
    warnings: list[str] = []
    cleaned_lines: list[str] = []
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
        match = RE_DEGENERATE_PLACEHOLDER_LINE.match(stripped)
        if match:
            warnings.append(f"已删除异常占位符行：{match.group('label')}")
            continue
        punctuation_count = sum(1 for ch in stripped if ch in "、,，")
        non_space_count = sum(1 for ch in stripped if not ch.isspace())
        if punctuation_count >= 20 and punctuation_count / max(non_space_count, 1) > 0.45:
            warnings.append("已删除一处异常标点重复内容")
            continue
        if RE_PLACEHOLDER_TEXT.search(stripped):
            warnings.append("已删除包含占位内容的段落")
            continue
        cleaned_lines.append(stripped)

    cleaned = "\n".join(cleaned_lines).strip()
    return cleaned, warnings


def _normalize_role_token(raw_role: str) -> str | None:
    role = str(raw_role or "").strip().lower().replace("-", "_")
    if role in ALLOWED_DRAFT_ROLES:
        return role
    return ROLE_LABEL_ALIASES.get(str(raw_role or "").strip())


def _match_role_line(line: str) -> tuple[str, str] | None:
    for pattern in (RE_ROLE_LINE, RE_ALT_ROLE_LINE):
        match = pattern.match(line)
        if match:
            role = _normalize_role_token(match.group("role"))
            if role:
                return role, match.group("text").strip()
    cn_match = RE_CN_ROLE_LINE.match(line)
    if cn_match:
        role = _normalize_role_token(cn_match.group("role"))
        if role:
            return role, cn_match.group("text").strip()
    return None


def _match_role_marker(line: str) -> str | None:
    for pattern in (RE_ROLE_MARKER, RE_ALT_ROLE_MARKER):
        match = pattern.match(line)
        if match:
            role = _normalize_role_token(match.group("role"))
            if role:
                return role
    cn_match = RE_CN_ROLE_MARKER.match(line)
    if cn_match:
        return _normalize_role_token(cn_match.group("role"))
    return None


def draft_has_role_markers(text: str) -> bool:
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line or line in ("```", "```text"):
            continue
        if _match_role_line(line) or _match_role_marker(line):
            return True
    return False


def normalize_draft_role_markers(text: str) -> tuple[str, list[str]]:
    """统一 role 标记写法，兼容单括号/中文标签等常见模型输出。"""
    warnings: list[str] = []
    normalized_lines: list[str] = []
    changed = False
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            normalized_lines.append("")
            continue
        matched = _match_role_line(line)
        if matched:
            role, value = matched
            if not RE_ROLE_LINE.match(line):
                changed = True
            normalized_lines.append(f"[[{role}]]{value}")
            continue
        marker = _match_role_marker(line)
        if marker:
            if not RE_ROLE_MARKER.match(line):
                changed = True
            normalized_lines.append(f"[[{marker}]]")
            continue
        normalized_lines.append(line)
    if changed:
        warnings.append("已将模型输出的段落标记统一为 [[role]] 格式")
    return "\n".join(normalized_lines).strip(), warnings


def coerce_plain_draft_to_structured(text: str) -> tuple[str, list[str]]:
    """模型未打 role 标记时，按公文常见版式推断段落角色。"""
    warnings = ["模型未返回 [[role]] 标记，已按公文版式自动分段"]
    blocks: list[dict[str, str]] = []
    paragraph_lines: list[str] = []
    title_assigned = False

    def flush_body() -> None:
        nonlocal paragraph_lines
        body = "\n".join(paragraph_lines).strip()
        if body:
            blocks.append({"role": "body", "text": body})
        paragraph_lines = []

    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line or line in ("```", "```text"):
            flush_body()
            continue
        if _match_role_line(line) or _match_role_marker(line):
            flush_body()
            matched = _match_role_line(line)
            if matched:
                blocks.append({"role": matched[0], "text": matched[1]})
            continue
        if RE_ATTACHMENT_START.match(line):
            flush_body()
            blocks.append({"role": "attachment_head", "text": line})
            continue
        if looks_like_parenthetical_subtitle(line) and not any(block["role"] == "subtitle" for block in blocks):
            # 版头括注单位，即使主标题尚未写出也先记为 subtitle
            flush_body()
            blocks.append({"role": "subtitle", "text": line})
            continue
        if RE_RECIPIENT_LINE.match(line):
            flush_body()
            blocks.append({"role": "recipient", "text": line})
            continue
        if re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", line):
            flush_body()
            blocks.append({"role": "sign_date", "text": line})
            continue
        if (
            re.search(r"(?:厅|局|部|委|办|院|队|中心|处|组|总队|支队|大队|政府|委员会)$", line)
            and 4 <= len(line) <= 40
            and title_assigned
        ):
            flush_body()
            blocks.append({"role": "sign_unit", "text": line})
            continue
        h2_match = RE_H2_TEXT.match(line)
        if h2_match:
            flush_body()
            blocks.append({"role": "h2", "text": line})
            continue
        h1_match = RE_H1_TEXT.match(line)
        if h1_match and len(h1_match.group("text")) <= 40:
            flush_body()
            blocks.append({"role": "h1", "text": line})
            continue
        if (
            not title_assigned
            and not looks_like_section_heading(line)
            and (
                RE_DOC_TITLE.match(line)
                or (len(line) <= 80 and line.endswith(("函", "通知", "意见", "报告", "请示", "通报", "材料", "汇报")))
                or ("会议" in line and 6 <= len(line) <= 40)
            )
        ):
            flush_body()
            blocks.append({"role": "title", "text": line})
            title_assigned = True
            continue
        paragraph_lines.append(line)
    flush_body()

    if not blocks:
        stripped = text.strip()
        if stripped:
            blocks.append({"role": "body", "text": stripped})
    if not blocks:
        return text, []
    # 仅当首段不像章节小标题时，才识别为主标题
    if (
        blocks[0]["role"] != "title"
        and len(blocks[0]["text"]) <= 80
        and not looks_like_section_heading(blocks[0]["text"])
        and blocks[0]["role"] not in {"h1", "h2", "h3", "h4", "recipient", "sign_unit", "sign_date"}
    ):
        blocks[0]["role"] = "title"
        warnings.append("已将首段识别为标题")
    structured = "\n".join(f"[[{block['role']}]]{block['text']}" for block in blocks if block.get("text"))
    return structured, warnings


def prepare_draft_for_parsing(text: str) -> tuple[str, list[str]]:
    cleaned, warnings = sanitize_draft_text(text)
    if not cleaned:
        return cleaned, warnings
    normalized, norm_warnings = normalize_draft_role_markers(cleaned)
    warnings.extend(norm_warnings)
    if draft_has_role_markers(normalized):
        return normalized, warnings
    coerced, coerce_warnings = coerce_plain_draft_to_structured(normalized)
    warnings.extend(coerce_warnings)
    return coerced, warnings


def call_document_drafting_model(
    request_url: str,
    api_key: str,
    model_name: str,
    messages: list[dict[str, str]],
    speed_mode: str,
    output_limit: int,
    temperature: float = 0.1,
) -> tuple[str, list[str], dict[str, Any]]:
    """调用模型按选定范文写法生成完整公文。"""
    budgets = DRAFT_TOKEN_BUDGETS.get(speed_mode, DRAFT_TOKEN_BUDGETS["standard"])
    attempts: list[dict[str, Any]] = []
    for max_tokens in budgets:
        result = call_chat_model_result(
            request_url, api_key, model_name, messages,
            temperature=temperature, max_tokens=max_tokens,
        )
        content = result.get("content") or ""
        usage = result.get("usage") or {}
        attempt = {
            "maxTokens": max_tokens,
            "finishReason": result.get("finishReason") or "unknown",
            "outputChars": len(content),
            "reasoningChars": result.get("reasoningChars", 0),
            "usage": usage,
        }
        attempts.append(attempt)
        if not content or RE_EMPTY_SUMMARY.search(content):
            if content:
                attempt["invalidOutput"] = "错误声明材料为空"
            continue
        prepared, prep_warnings = prepare_draft_for_parsing(content)
        if not prepared or not draft_has_role_markers(prepared):
            attempt["invalidOutput"] = "未识别到可用段落结构"
            continue
        try:
            parse_structured_draft(prepared)
        except RuntimeError as exc:
            attempt["invalidOutput"] = str(exc)
            continue
        if len(prepared) > output_limit:
            raise RuntimeError(
                f"模型成文结果为 {len(prepared)} 字，超过 {output_limit} 字上限；"
                "已停止生成以避免后端截断。"
            )
        return prepared, prep_warnings, {
            "finishReason": attempt["finishReason"],
            "attempts": attempts,
            "outputChars": len(prepared),
            "usage": usage,
            "normalizedFromPlainText": bool(prep_warnings),
        }
    details = "；".join(
        f"第 {index + 1} 次 max_tokens={item['maxTokens']}，finish_reason={item['finishReason']}，"
        f"正文={item['outputChars']} 字，推理={item['usage'].get('completion_tokens_details', {}).get('reasoning_tokens', 0)} tokens"
        f"{('，' + item['invalidOutput']) if item.get('invalidOutput') else ''}"
        for index, item in enumerate(attempts)
    )
    raise RuntimeError(f"模型成文未返回正文：{details}。请检查模型服务或改用可返回文本的模型。")


def call_transformation_planner(
    request_url: str, api_key: str, model_name: str,
    messages: list[dict[str, str]], speed_mode: str,
    materials: list[dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, Any]]:
    budgets = DRAFT_TOKEN_BUDGETS.get(speed_mode, DRAFT_TOKEN_BUDGETS["standard"])
    max_tokens = min(max(budgets), 8192)
    result = call_chat_model_result(
        request_url, api_key, model_name, messages,
        temperature=0.0, max_tokens=max_tokens,
    )
    content = result.get("content") or ""
    if not content:
        raise RuntimeError("范文结构分析模型没有返回内容")
    meta = {
        "maxTokens": max_tokens,
        "outputChars": len(content),
        "usage": result.get("usage") or {},
    }
    try:
        plan = parse_transformation_plan(content, materials)
    except RuntimeError as exc:
        source_ids = _source_ids(materials)
        plan = {
            "documentPurpose": "按目标文种形成正式公文",
            "organization": {
                "opening": "依据用户要求和材料事实组织开头",
                "development": "按事项逻辑展开主体内容",
                "closing": "按目标文种规范收束",
                "style": ["正式", "准确", "简洁"],
            },
            "targetSections": [
                {
                    "order": 1,
                    "role": "body",
                    "function": "承载上传材料中的核心事实和工作要求",
                    "writingMethod": "仅依据材料事实组织，不引入范文事实",
                    "sourceIds": source_ids,
                }
            ],
        }
        meta["fallback"] = True
        meta["fallbackReason"] = str(exc)
    return plan, meta


def validate_reference_fact_isolation(
    draft: str, reference: dict[str, Any], prompt: str, materials: list[dict[str, Any]],
    strict: bool | None = None,
) -> dict[str, Any]:
    """审计范文事实隔离；strict=True 时检测到泄漏则阻断成文。"""
    source_text = "\n".join([prompt, *(item.get("text", "") for item in materials)])
    source_facts = set(RE_REFERENCE_FACT_TOKEN.findall(source_text))
    reference_facts = set(RE_REFERENCE_FACT_TOKEN.findall(reference.get("text", "")))
    output_facts = set(RE_REFERENCE_FACT_TOKEN.findall(draft))
    leaked = sorted((reference_facts - source_facts) & output_facts, key=lambda item: (-len(item), item))
    warnings = []
    if leaked:
        warnings.append(f"检测到参考范文独有数字或日期，需人工核验：{'、'.join(leaked[:8])}")

    normalize = lambda value: "".join(re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", value))
    source_normalized = normalize(source_text)
    reference_normalized = normalize(reference.get("text", ""))
    draft_normalized = normalize(draft)
    # 短公文术语（如“推进风险预警”“行业监管责任”）会在多个文种中重复出现，
    # 不能仅凭一个短片段就认定范文事实串入。这里只拦截多个长连续片段的直接复用。
    phrase_size = 16
    action_signal = re.compile(
        r"务必|必须|应当|制定|方案|预案|监督|组织领导|严格按照|确保|落实|建立|完善|推进|责任"
    )
    reference_phrases = {
        reference_normalized[index:index + phrase_size]
        for index in range(max(0, len(reference_normalized) - phrase_size + 1))
    }
    leaked_phrases = sorted(
        phrase for phrase in reference_phrases
        if (phrase not in source_normalized and phrase in draft_normalized
            and action_signal.search(phrase))
    )
    if len(leaked_phrases) >= 2:
        warnings.append(
            f"检测到可能复用参考范文长句，需人工核验：{'、'.join(leaked_phrases[:3])}"
        )
    strict = (
        strict
        if strict is not None
        else os.getenv("AI_STRICT_REFERENCE_ISOLATION", "").strip().lower() in {"1", "true", "yes", "on"}
    )
    isolation_risk = bool(strict and (leaked or len(leaked_phrases) >= 2))
    if isolation_risk:
        warnings.append("检测到范文表述串入风险，将触发结构化仿写重构")
    return {
        "referenceFactTokens": leaked[:8],
        "referencePhraseMatches": len(leaked_phrases),
        "isolationRisk": isolation_risk,
        "warnings": warnings,
    }


def _normalized_shingles(text: str, size: int = 8) -> set[str]:
    normalized = "".join(re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", text))
    return {
        normalized[index:index + size]
        for index in range(max(0, len(normalized) - size + 1))
    }


def _lcs_length(left: list[str], right: list[str]) -> int:
    previous = [0] * (len(right) + 1)
    for left_item in left:
        current = [0]
        for index, right_item in enumerate(right, start=1):
            current.append(previous[index - 1] + 1 if left_item == right_item else max(previous[index], current[-1]))
        previous = current
    return previous[-1]


def validate_transformed_draft(
    blocks: list[dict[str, str]], materials: list[dict[str, Any]],
    transformation_plan: dict[str, Any],
) -> dict[str, Any]:
    """评估成文质量；相似度偏高时返回 similarityRisk 供重构循环处理，不再硬拦截。"""
    draft = "\n".join(block.get("text", "") for block in blocks).strip()
    if RE_PLACEHOLDER_TEXT.search(draft):
        raise RuntimeError("生成结果仍包含未填写占位符，已停止输出；请在上传材料中提供对应事实")

    source = "\n".join(
        item.get("rawText") or item.get("text", "") for item in materials
    )
    brief_source = "\n".join(item.get("text", "") for item in materials)
    source_shingles = _normalized_shingles(source)
    brief_shingles = _normalized_shingles(brief_source)
    draft_shingles = _normalized_shingles(draft)
    source_copy_coverage = (
        len(source_shingles & draft_shingles) / len(draft_shingles)
        if draft_shingles else 0.0
    )
    brief_copy_coverage = (
        len(brief_shingles & draft_shingles) / len(draft_shingles)
        if draft_shingles and brief_shingles else source_copy_coverage
    )
    copied_chars = 0
    source_paragraphs = [
        part.strip()
        for part in re.split(r"\n\s*\n", brief_source or source)
        if len(part.strip()) >= 40
    ]
    for block in blocks:
        text = block.get("text", "").strip()
        if len(text) < 40:
            continue
        best_ratio = max(
            (SequenceMatcher(None, text, paragraph).ratio() for paragraph in source_paragraphs),
            default=0.0,
        )
        if best_ratio >= 0.9:
            copied_chars += len(text)
    copied_ratio = copied_chars / max(len(draft), 1)
    expected_roles = [
        str(section.get("role"))
        for section in transformation_plan.get("targetSections", [])
        if section.get("role") != "other"
    ]
    actual_roles = [block.get("role") for block in blocks if block.get("role") != "other"]
    structure_coverage = _lcs_length(expected_roles, actual_roles) / max(len(expected_roles), 1)
    warnings = []
    similarity_risk = (
        source_copy_coverage >= SIMILARITY_REWRITE_COPY_THRESHOLD
        or brief_copy_coverage >= SIMILARITY_REWRITE_COPY_THRESHOLD
        or copied_ratio >= SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD
    )
    if similarity_risk:
        warnings.append(
            "生成结果与材料原文重合度偏高，将自动进入模板框架仿写重构"
        )
    if structure_coverage < 0.6:
        warnings.append("成文角色序列与参考范文差异较大，建议人工核验段落组织")
    return {
        "sourceCopyCoverage": round(source_copy_coverage, 4),
        "briefCopyCoverage": round(brief_copy_coverage, 4),
        "copiedParagraphRatio": round(copied_ratio, 4),
        "structureCoverage": round(structure_coverage, 4),
        "similarityRisk": similarity_risk,
        "warnings": warnings,
    }


def parse_structured_draft(text: str) -> tuple[list[dict[str, str]], list[str]]:
    blocks: list[dict[str, str]] = []
    warnings: list[str] = []
    pending_role: str | None = None
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line or line in ("```", "```text"):
            continue
        matched = _match_role_line(line)
        marker_role = _match_role_marker(line)
        if marker_role:
            pending_role = marker_role
            continue
        if matched:
            role, value = matched
            if value and not RE_PLACEHOLDER_TEXT.search(value):
                blocks.append({"role": role, "text": value})
            elif value:
                warnings.append("已删除包含占位内容的段落")
            pending_role = None
            continue
        if not matched:
            if pending_role:
                role = pending_role if pending_role in ALLOWED_DRAFT_ROLES else "body"
                blocks.append({"role": role, "text": line})
                pending_role = None
                continue
            if not blocks:
                raise RuntimeError(f"模型首段未按结构化段落格式输出：{line[:60]}")
            if blocks[-1]["role"] in {"h1", "h2", "h3", "h4"}:
                blocks.append({"role": "body", "text": line})
                warnings.append("已将层级标题后的未标注续行转换为独立正文段")
            else:
                blocks[-1]["text"] = f"{blocks[-1]['text']}\n{line}"
                warnings.append("已合并一处未标注角色的续行")
            continue
    if pending_role:
        warnings.append(f"模型返回了无正文的段落角色 {pending_role}，已忽略")
    if not blocks:
        raise RuntimeError("模型没有返回可用公文段落")
    return blocks, warnings


def validate_document_blocks(blocks: list[dict[str, str]], template_id: str) -> list[str]:
    roles = {block["role"] for block in blocks}
    rules = get_template_classification(template_id)
    warnings: list[str] = []
    if "title" not in roles:
        warnings.append("生成结果缺少标题，将触发定向补齐")
    if rules.get("recipient") == "required" and "recipient" not in roles:
        warnings.append("该文种通常需要主送机关，但用户材料未提供可用内容，未自动补写")
    if rules.get("signature") == "required":
        missing = [role for role in ("sign_unit", "sign_date") if role not in roles]
        if missing:
            warnings.append("该文种通常需要落款信息，但用户材料未提供，未自动补写")
    return warnings


def detect_missing_required_fields(
    blocks: list[dict[str, str]],
    reference: dict[str, Any],
) -> list[dict[str, Any]]:
    """检测仿写结果相对模板必备字段的缺失项。"""
    present = {block.get("role") for block in blocks}
    blueprint = reference.get("structureJson") or {}
    required = list(blueprint.get("requiredRoles") or ["title"])
    if "title" not in required:
        required.insert(0, "title")
    section_map = {
        str(section.get("role")): section
        for section in (blueprint.get("sections") or [])
        if section.get("role")
    }
    missing = []
    for role in required:
        if role in present:
            continue
        section = section_map.get(role) or {}
        missing.append({
            "role": role,
            "label": ROLE_NAMES.get(role, role),
            "function": section.get("function") or _default_section_function(role),
            "writingMethod": section.get("writingMethod") or _default_writing_method(role),
            "formatRules": section.get("formatRules") or _default_format_rules(role),
            "constraintTags": section.get("constraintTags") or _default_constraint_tags(role),
            "reason": f"缺少{ROLE_NAMES.get(role, role)}",
        })
    return missing


def derive_title_from_materials(
    materials: list[dict[str, Any]],
    prompt: str,
    template_label: str,
) -> str:
    """规则侧从材料归纳标题，避免缺失时直接失败。"""
    candidates: list[str] = []
    for material in materials:
        prepared_title = str(material.get("title") or "").strip()
        if prepared_title and not looks_like_section_heading(prepared_title):
            return prepared_title[:40]
        for element in material.get("sourceElements") or []:
            if element.get("role") == "title" and element.get("text"):
                text = str(element["text"]).strip()
                if not looks_like_section_heading(text):
                    return text[:40]
        zone_title = str((material.get("materialZones") or {}).get("titleText") or "").strip()
        if zone_title and not looks_like_section_heading(zone_title):
            candidates.append(zone_title[:40])
        summary = str(material.get("summary") or "").strip()
        if summary:
            first = summary.split("。")[0][:36]
            if not looks_like_section_heading(first):
                candidates.append(first)
        for fact in material.get("structuredFacts") or []:
            content = str(fact.get("content") or "").strip()
            if (
                8 <= len(content) <= 40
                and not content.endswith(("：", ":"))
                and not looks_like_section_heading(content)
            ):
                candidates.append(content)
                break
        raw = str(material.get("rawText") or material.get("text") or "")
        for line in raw.splitlines():
            value = line.strip()
            if not value or value.startswith("["):
                continue
            if looks_like_section_heading(value) or looks_like_parenthetical_subtitle(value):
                continue
            if 6 <= len(value) <= 40 and not re.search(r"[。；;]$", value):
                candidates.append(value)
                break
    prompt_title = (prompt or "").strip().splitlines()[0] if prompt else ""
    if 6 <= len(prompt_title) <= 40 and not looks_like_section_heading(prompt_title):
        candidates.append(prompt_title)
    for item in candidates:
        cleaned = re.sub(r"^(关于|有关)", "", item).strip(" 。；;，,")
        if cleaned and not looks_like_section_heading(cleaned):
            if "会议" in cleaned or "材料" in cleaned or "调度" in cleaned:
                return cleaned[:40]
            if cleaned.endswith(("函", "通知", "意见", "报告", "汇报")):
                return cleaned[:40]
            if not cleaned.startswith("关于") and template_label:
                return f"关于{cleaned}"[:40]
            return cleaned[:40]
    label = template_label or "公文"
    return f"关于有关事项的{label}"[:40]


def repair_promoted_section_title(blocks: list[dict[str, str]]) -> tuple[list[dict[str, str]], list[str]]:
    """把误当作主标题的「一、…」恢复为一级标题，供后续补齐真正主标题。"""
    if not blocks:
        return blocks, []
    result = [dict(block) for block in blocks]
    warnings: list[str] = []
    title_idx = next((i for i, block in enumerate(result) if block.get("role") == "title"), None)
    if title_idx is None:
        return result, warnings
    title_text = str(result[title_idx].get("text") or "").strip()
    if not looks_like_section_heading(title_text):
        return result, warnings
    result[title_idx]["role"] = "h1"
    warnings.append("已将误作主标题的一级小标题恢复为 [[h1]]")
    next_idx = title_idx + 1
    if next_idx < len(result) and looks_like_parenthetical_subtitle(result[next_idx].get("text") or ""):
        result[next_idx]["role"] = "subtitle"
    return result, warnings


def normalize_document_header_order(blocks: list[dict[str, str]]) -> list[dict[str, str]]:
    """保证版头顺序：title → subtitle → recipient → security → 其余。"""
    if not blocks:
        return blocks
    header_roles = ("title", "subtitle", "recipient", "security")
    headers = {role: [] for role in header_roles}
    rest: list[dict[str, str]] = []
    for block in blocks:
        role = block.get("role")
        if role in headers and not headers[role]:
            headers[role].append(dict(block))
        elif role in headers:
            # 多余同角色版头并入 rest，避免重复标题
            rest.append(dict(block))
        else:
            rest.append(dict(block))
    ordered = []
    for role in header_roles:
        ordered.extend(headers[role])
    ordered.extend(rest)
    return ordered


def _insert_filled_block(
    blocks: list[dict[str, str]],
    role: str,
    text: str,
) -> list[dict[str, str]]:
    result = [dict(block) for block in blocks]
    if any(block.get("role") == role for block in result):
        return result
    if role == "title":
        result.insert(0, {"role": role, "text": text})
        return result
    if role in {"subtitle", "recipient", "security"}:
        header_order = ["title", "subtitle", "recipient", "security"]
        insert_at = 0
        for index, block in enumerate(result):
            current = block.get("role")
            if current in header_order and header_order.index(current) < header_order.index(role):
                insert_at = index + 1
            elif current not in header_order:
                break
            else:
                insert_at = index + 1
        result.insert(insert_at, {"role": role, "text": text})
        return result
    result.append({"role": role, "text": text})
    return result


def fill_missing_fields_targeted(
    blocks: list[dict[str, str]],
    missing_fields: list[dict[str, Any]],
    materials: list[dict[str, Any]],
    prompt: str,
    template: dict[str, Any],
    reference: dict[str, Any],
    request_url: str,
    api_key: str,
    model_name: str,
) -> tuple[list[dict[str, str]], list[str], dict[str, Any]]:
    """权限隔离的定向补齐：只补缺失字段，不重写已有合规段落。"""
    if not missing_fields:
        return blocks, [], {"filledRoles": [], "mode": "noop"}

    warnings: list[str] = []
    result = [dict(block) for block in blocks]
    filled_roles: list[str] = []
    template_label = template.get("label") or template.get("name") or "公文"
    structure_payload = reference_structure_payload(reference)
    material_brief = "\n".join(
        item.get("summary") or item.get("text", "")[:1200]
        for item in materials
    ).strip()
    facts = [
        f"{fact.get('id')}:{fact.get('content')}"
        for item in materials
        for fact in (item.get("structuredFacts") or [])[:20]
        if fact.get("content")
    ]

    # 先规则补标题，尽量少调模型
    remaining = []
    for field in missing_fields:
        role = field["role"]
        if role == "title":
            title = derive_title_from_materials(materials, prompt, template_label)
            result = _insert_filled_block(result, "title", title)
            filled_roles.append("title")
            warnings.append(f"已定向补齐大标题：{title}")
        else:
            remaining.append(field)

    if remaining:
        missing_spec = json.dumps(remaining, ensure_ascii=False, indent=2)
        messages = [
            {
                "role": "system",
                "content": (
                    "你是公文缺失字段定向补齐器。权限严格受限："
                    "1) 只生成指定缺失字段；2) 禁止重写、合并或改写已有段落；"
                    "3) 禁止二次全量解析全文；4) 不得引入材料之外的事实；"
                    "5) 只输出 JSON。"
                ),
            },
            {
                "role": "user",
                "content": (
                    f"目标文种：{template_label}\n"
                    f"用户要求：{prompt or '形成正式公文'}\n"
                    f"模板可调用结构：{json.dumps({k: structure_payload.get(k) for k in ('requiredRoles', 'organization', 'rhetoricPatterns')}, ensure_ascii=False)}\n"
                    f"缺失字段规格：\n{missing_spec}\n\n"
                    f"材料事实摘要：\n{material_brief[:3000]}\n"
                    f"事实条目：\n{chr(10).join(facts[:20]) or '无'}\n\n"
                    '返回 JSON：{"fills":[{"role":"title","text":"..."}]}\n'
                    "fills 只能包含上述缺失 role；text 须可由材料支撑。"
                ),
            },
        ]
        try:
            response = call_chat_model_result(
                request_url, api_key, model_name, messages, temperature=0.1, max_tokens=1024,
            )
            payload = _parse_json_object(response.get("content") or "")
            allowed = {item["role"] for item in remaining}
            for item in payload.get("fills") or []:
                role = str(item.get("role") or "")
                text = str(item.get("text") or "").strip()
                if role not in allowed or not text or RE_PLACEHOLDER_TEXT.search(text):
                    continue
                result = _insert_filled_block(result, role, text)
                filled_roles.append(role)
                warnings.append(f"已定向补齐{ROLE_NAMES.get(role, role)}")
        except Exception as exc:
            warnings.append(f"定向补齐部分失败：{exc}")

    # 标题兜底：AI/规则后仍缺则再补一次规则标题
    if not any(block.get("role") == "title" for block in result):
        title = derive_title_from_materials(materials, prompt, template_label)
        result = _insert_filled_block(result, "title", title)
        filled_roles.append("title")
        warnings.append(f"已规则兜底补齐大标题：{title}")

    return result, warnings, {
        "filledRoles": filled_roles,
        "requestedRoles": [item["role"] for item in missing_fields],
        "mode": "targeted_fill",
        "permissionIsolation": True,
    }


def normalize_generated_block_roles(
    blocks: list[dict[str, str]], template_id: str,
) -> list[str]:
    """纠正模型把正文分项标识误报为标题层级的情况。"""
    config = get_document_config(template_id)
    corrected = 0
    for block in blocks:
        role = block["role"]
        normalized_role = normalize_role_for_text(role, block["text"], config)
        if normalized_role != role:
            block["role"] = normalized_role
            corrected += 1
    if not corrected:
        return []
    return [f"已将 {corrected} 个‘一是、二是’类正文分项从标题层级纠正为正文"]


def validate_heading_skeleton(blocks: list[dict[str, str]], reference: dict[str, Any]) -> list[str]:
    """轻量提示标题层级是否偏少，不推动模型凑标题数量。"""
    blueprint = reference.get("structureJson") or {}
    expected = list(blueprint.get("headingSkeleton") or [])
    if not expected:
        return []
    actual = [block["role"] for block in blocks if block["role"] in {"h1", "h2", "h3", "h4"}]
    if not actual and any(block["role"] == "body" for block in blocks):
        return ["主体以连续正文展开，未设层级标题；如材料块较多可人工考虑是否加分段标题"]
    return []


def preserve_source_elements(
    blocks: list[dict[str, str]], materials: list[dict[str, Any]],
) -> tuple[list[dict[str, str]], list[str]]:
    """以上传材料为准补回模型遗漏或改写的公文要素。"""
    result = [dict(block) for block in blocks]
    warnings: list[str] = []
    required = [
        element
        for material in materials
        for element in material.get("sourceElements", [])
    ]
    for element in required:
        role, text = element["role"], element["text"]
        if any(block.get("role") == role and block.get("text") == text for block in result):
            continue
        if role == "title":
            if looks_like_section_heading(text):
                warnings.append("材料标题疑似一级小标题，已跳过强制保留，改由成文补齐主标题")
                continue
            if result and result[0].get("role") == "title":
                if looks_like_section_heading(result[0].get("text") or ""):
                    result[0]["text"] = text
                else:
                    result[0]["text"] = text
            else:
                result.insert(0, {"role": role, "text": text})
            warnings.append("已按上传材料保留标题")
            continue
        if role == "subtitle":
            if any(block.get("role") == "subtitle" for block in result):
                continue
            insert_at = 1 if result and result[0].get("role") == "title" else 0
            result.insert(insert_at, {"role": role, "text": text})
            warnings.append("已按上传材料保留副标题")
            continue
        if role == "sign_unit":
            result = [block for block in result if block.get("role") != "sign_unit"]
            units: list[str] = []
            for material in materials:
                for unit in (material.get("materialZones") or {}).get("signUnits") or []:
                    if unit.strip() and unit.strip() not in units:
                        units.append(unit.strip())
            if not units:
                units = [part.strip() for part in text.splitlines() if part.strip()]
            insert_at = next(
                (index for index, block in enumerate(result) if block.get("role") == "sign_date"),
                len(result),
            )
            for offset, unit in enumerate(units):
                result.insert(insert_at + offset, {"role": "sign_unit", "text": unit})
            warnings.append("已按上传材料保留落款单位（逐行）")
            continue
        same_role = [index for index, block in enumerate(result) if block.get("role") == role]
        if same_role:
            result[same_role[0]]["text"] = text
        elif role in {"subtitle", "recipient", "security"}:
            header_roles = {"title", "subtitle", "recipient", "security"}
            insert_at = sum(1 for block in result if block.get("role") in header_roles)
            result.insert(insert_at, {"role": role, "text": text})
        else:
            result.append({"role": role, "text": text})
        warnings.append(f"已按上传材料保留{ROLE_NAMES.get(role, role)}")
    return result, warnings


def expand_multiline_role_blocks(blocks: list[dict[str, str]]) -> list[dict[str, str]]:
    """标题、落款等多行要素拆成同 role 的多个段落，便于套格式时分行。"""
    expanded: list[dict[str, str]] = []
    for block in blocks:
        role = block.get("role", "body")
        text = str(block.get("text") or "").strip()
        if role in {"title", "sign_unit", "sign_contact"} and "\n" in text:
            for line in [part.strip() for part in text.splitlines() if part.strip()]:
                expanded.append({"role": role, "text": line})
        else:
            expanded.append(dict(block))
    return expanded


def split_generated_heading_blocks(
    blocks: list[dict[str, str]],
) -> tuple[list[dict[str, str]], list[str]]:
    """将模型误写在层级标题后的正文拆为独立正文段，避免整段套标题样式。"""
    result: list[dict[str, str]] = []
    warnings: list[str] = []
    for block in blocks:
        role = block.get("role", "body")
        text = str(block.get("text") or "").strip()
        if role not in {"h1", "h2", "h3", "h4"}:
            result.append(dict(block))
            continue
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if len(lines) < 2:
            result.append(dict(block))
            continue
        heading = lines[0].rstrip("，,；;：:。！？ ")
        body = "\n".join(lines[1:]).strip()
        if len(heading) < 3 or len(body) < 4:
            result.append(dict(block))
            continue
        result.append({"role": role, "text": heading})
        result.append({"role": "body", "text": body})
        warnings.append(f"已将{ROLE_NAMES.get(role, role)}后的正文拆分为独立段落")
    return result, warnings


def write_structured_docx(blocks: list[dict[str, str]], output_path: Path) -> dict[int, str]:
    doc = Document()
    roles: dict[int, str] = {}
    paragraph_index = 0
    for block in expand_multiline_role_blocks(blocks):
        text = str(block.get("text") or "").strip()
        if not text:
            continue
        doc.add_paragraph(text)
        roles[paragraph_index] = block["role"]
        paragraph_index += 1
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return roles


def generate_document(
    prompt: str,
    upload_paths: list[Path],
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str | None = None,
    temperature: float = 0.2,
) -> dict[str, Any]:
    return generate_documents(
        prompt=prompt,
        upload_paths=upload_paths,
        request_url=request_url,
        api_key=api_key,
        model_name=model_name,
        template_id=template_id,
        temperature=temperature,
    )["documents"][0]


def _generate_one_document(
    prompt: str,
    material: dict[str, Any],
    index: int,
    templates: list[dict[str, Any]],
    template_id: str | None,
    request_url: str,
    api_key: str,
    model_name: str,
    temperature: float,
    speed_mode: str,
    strict_reference_isolation: bool = False,
    format_output: bool = True,
) -> dict[str, Any]:
    resolved_template_id = choose_template_id(prompt, [material["text"]], template_id)
    template = next((item for item in templates if item["id"] == resolved_template_id), None)
    if not template:
        raise RuntimeError(f"未找到目标公文模板：{resolved_template_id}")
    warnings: list[str] = list(material.get("warnings") or [])
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
            "sourceElements": [],
        }]
    else:
        prepared_material, material_warnings = prepare_material(
            path=material["path"],
            prompt=prompt,
            template_label=template_label,
            request_url=request_url,
            api_key=api_key,
            model_name=model_name,
            template_id=resolved_template_id,
            inventory=material["inventory"],
        )
        current_materials = [prepared_material]
        warnings.extend(material_warnings)
    reference = select_reference_for_material(template, prompt, current_materials)
    if reference.get("warning"):
        warnings.append(f"参考范文解析提示：{reference['warning']}")
    plan_started_at = perf_counter()
    plan_messages = build_transformation_plan_messages(
        prompt, current_materials, template, reference,
    )
    transformation_plan, planner_meta = call_transformation_planner(
        request_url, api_key, model_name, plan_messages, speed_mode, current_materials,
    )
    plan_seconds = round(perf_counter() - plan_started_at, 2)
    messages = build_ai_messages(
        prompt, current_materials, template, reference, transformation_plan,
    )
    drafting_started_at = perf_counter()
    rewrite_retries = 0
    similarity_rewrites = 0
    targeted_fill_meta: dict[str, Any] = {"filledRoles": [], "mode": "noop"}
    validation_metrics: dict[str, Any] = {}
    last_error: RuntimeError | None = None
    last_reference_audit: dict[str, Any] | None = None
    draft = ""
    generation_meta: dict[str, Any] = {}
    blocks: list[dict[str, str]] = []
    base_messages = list(messages)
    for guard_attempt in range(MAX_GENERATION_GUARD_ATTEMPTS):
        if last_error is not None:
            if is_similarity_failure(last_error):
                messages = build_similarity_rewrite_messages(
                    prompt, current_materials, template, reference, transformation_plan,
                    last_error,
                    reference_audit=last_reference_audit,
                    validation_metrics=validation_metrics,
                    previous_draft=draft,
                )
                similarity_rewrites += 1
                draft_temperature = min(max(temperature, 0.25), 0.4)
            else:
                messages = [
                    *messages,
                    {"role": "assistant", "content": draft},
                    {"role": "user", "content": build_parse_retry_message(last_error)},
                ]
                draft_temperature = min(temperature, 0.2)
        else:
            messages = base_messages
            draft_temperature = min(temperature, 0.2)
        draft, draft_warnings, generation_meta = call_document_drafting_model(
            request_url, api_key, model_name, messages, speed_mode,
            output_limit=draft_char_limit(current_materials, reference), temperature=draft_temperature,
        )
        warnings.extend(draft_warnings)
        try:
            reference_audit = validate_reference_fact_isolation(
                draft, reference, prompt, current_materials, strict=strict_reference_isolation,
            )
            last_reference_audit = reference_audit
            warnings.extend(reference_audit["warnings"])
            blocks, structure_warnings = parse_structured_draft(draft)
            blocks, title_repair_warnings = repair_promoted_section_title(blocks)
            blocks, preservation_warnings = preserve_source_elements(blocks, current_materials)
            blocks, title_repair_warnings_2 = repair_promoted_section_title(blocks)
            blocks = expand_multiline_role_blocks(blocks)
            role_warnings = normalize_generated_block_roles(blocks, resolved_template_id)
            blocks, heading_split_warnings = split_generated_heading_blocks(blocks)
            document_warnings = validate_document_blocks(blocks, resolved_template_id)
            heading_warnings = validate_heading_skeleton(blocks, reference)
            warnings.extend(structure_warnings)
            warnings.extend(title_repair_warnings)
            warnings.extend(preservation_warnings)
            warnings.extend(title_repair_warnings_2)
            warnings.extend(role_warnings)
            warnings.extend(heading_split_warnings)
            warnings.extend(document_warnings)
            warnings.extend(heading_warnings)

            missing_fields = detect_missing_required_fields(blocks, reference)
            if missing_fields:
                blocks, fill_warnings, targeted_fill_meta = fill_missing_fields_targeted(
                    blocks, missing_fields, current_materials, prompt, template, reference,
                    request_url, api_key, model_name,
                )
                warnings.extend(fill_warnings)
                still_missing = detect_missing_required_fields(blocks, reference)
                if still_missing and any(item["role"] == "title" for item in still_missing):
                    raise RuntimeError("定向补齐后仍缺少目标公文大标题")
            blocks = normalize_document_header_order(blocks)

            validation_metrics = validate_transformed_draft(
                blocks, current_materials, transformation_plan,
            )
            validation_metrics["referenceIsolation"] = reference_audit
            validation_metrics["targetedFill"] = targeted_fill_meta
            warnings.extend(validation_metrics.get("warnings") or [])
            similarity_risk = bool(
                reference_audit.get("isolationRisk") or validation_metrics.get("similarityRisk")
            )
            if similarity_risk:
                reason_parts = []
                if reference_audit.get("isolationRisk"):
                    reason_parts.append("范文表述串入")
                if validation_metrics.get("similarityRisk"):
                    reason_parts.append("材料原文重合偏高")
                if guard_attempt >= MAX_GENERATION_GUARD_ATTEMPTS - 1:
                    warnings.append(
                        "已达最大框架仿写重构次数，保留当前版本；建议人工核验相似度与标题层级"
                    )
                    break
                rewrite_retries += 1
                last_error = RuntimeError(
                    f"相似度偏高，自动进入模板结构化仿写重构：{'、'.join(reason_parts)}"
                )
                warnings.append(
                    f"相似度风险触发，自动进入范文结构化仿写重构（第 {similarity_rewrites + 1} 次）：{last_error}"
                )
                continue
            if similarity_rewrites:
                warnings.append(
                    f"已通过语义差异化重构 {similarity_rewrites} 次，规避范文/材料同质化风险"
                )
            last_error = None
            break
        except RuntimeError as exc:
            if guard_attempt >= MAX_GENERATION_GUARD_ATTEMPTS - 1:
                raise
            rewrite_retries += 1
            last_error = exc
            if is_similarity_failure(exc):
                warnings.append(
                    f"相似度风控触发，自动进入范文结构化仿写重构（第 {similarity_rewrites + 1} 次）：{exc}"
                )
            else:
                warnings.append(f"结构化输出无法使用，已自动重试：{exc}")
    generation_meta["rewriteRetries"] = rewrite_retries
    generation_meta["similarityRewrites"] = similarity_rewrites
    generation_meta["targetedFill"] = targeted_fill_meta
    generation_meta["planner"] = planner_meta
    generation_meta["planSeconds"] = plan_seconds
    generation_meta["validation"] = validation_metrics

    # 写作阶段只负责生成结构化原稿；兼容调用可继续在此完成格式转换。
    raw_id = uuid.uuid4().hex
    raw_path = TEMP_DIR / f"{raw_id}_ai_raw.docx"
    result_id = uuid.uuid4().hex
    result_path = TEMP_DIR / f"{result_id}_ai.docx"
    roles = write_structured_docx(blocks, raw_path)
    output_path = raw_path
    paragraphs = [{"index": index, "text": block["text"], "role": block["role"],
                   "confidence": 1.0, "lowConfidence": False} for index, block in enumerate(blocks)]
    if format_output:
        convert_with_roles(
            str(raw_path), str(result_path), roles=roles,
            template_id=resolved_template_id, split_embedded_markers=False,
            split_inline_headings=False, force_heading_bold=True,
        )
        output_path = result_path
        paragraphs = build_structure(str(result_path), resolved_template_id, split_embedded_markers=False)
        for paragraph in paragraphs:
            role = roles.get(paragraph["index"])
            if role:
                paragraph["role"] = role
                paragraph["confidence"] = 1.0
                paragraph["lowConfidence"] = False
    preview = "\n".join(block["text"] for block in blocks)
    stem = Path(material["name"]).stem or f"document_{index}"
    return {
        "documentId": result_id if format_output else raw_id,
        "path": str(output_path),
        "rawPath": str(raw_path),
        "roles": roles,
        "formatted": format_output,
        "filename": f"{stem}_AI公文.docx",
        "sourceName": material["name"],
        "templateId": resolved_template_id,
        "templateName": template["name"] if template else "",
        "reference": {
            "filename": reference["filename"],
            "similarity": reference["similarity"],
            "textChars": reference["textChars"],
        },
        "transformationPlan": transformation_plan,
        "preview": preview[:3000],
        "paragraphs": paragraphs,
        "warnings": warnings,
        "processingMode": current_materials[0].get("processingMode", "direct"),
        "fullTextLength": current_materials[0].get("fullTextLength", 0),
        "chunkCount": current_materials[0].get("chunkCount", 0),
        "readReport": current_materials[0].get("readReport"),
        "speedMode": speed_mode,
        "generation": generation_meta,
        "draftLength": len(draft),
        "summaryLength": len(draft),
        "stages": {
            "extractionSeconds": material.get("extractionSeconds", 0),
            "draftingSeconds": round(perf_counter() - drafting_started_at, 2),
            "rendering": "材料成文 + 体例坐标 + 后端模板规则",
        },
    }


def generate_documents(
    prompt: str,
    upload_paths: list[Path],
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str | None = None,
    temperature: float = 0.2,
    speed_mode: str = "standard",
    strict_reference_isolation: bool = False,
    format_output: bool = True,
) -> dict[str, Any]:
    started_at = perf_counter()
    if speed_mode not in {"fast", "standard", "deep"}:
        speed_mode = "standard"
    if not template_id or normalize_template_id(template_id) == "generic":
        raise RuntimeError("请选择目标公文模板（通知、报告、函等）后再生成")
    if not upload_paths:
        raise RuntimeError("请至少上传一份材料后再生成，避免无依据空写")
    templates = scan_template_library()
    documents: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    targets = []
    for path in upload_paths:
        extraction_started_at = perf_counter()
        name = display_name_from_temp_path(path)
        try:
            inventory = extract_document_inventory(path)
            text = inventory_to_text(inventory).strip()
            extract_warnings = list(inventory.get("warnings") or [])
        except Exception as exc:
            failures.append({"filename": name, "reason": f"文本提取失败：{exc}"})
            continue
        if not text.strip():
            detail = "；".join(extract_warnings) or "未提取到可用文本"
            failures.append({"filename": name, "reason": f"无法处理：{detail}"})
            continue
        material_quality = assess_material_content(inventory)
        if material_quality["status"] != "ready":
            failures.append({
                "filename": name,
                "reason": f"材料正文不足：{material_quality['reason']}。请上传正文页、附件或完整文件",
            })
            continue
        targets.append({
            "path": path,
            "name": name,
            "text": text,
            "warnings": extract_warnings,
            "inventory": inventory,
            "extractionSeconds": round(perf_counter() - extraction_started_at, 2),
        })

    for index, material in enumerate(targets, start=1):
        try:
            documents.append(_generate_one_document(
                prompt=prompt,
                material=material,
                index=index,
                templates=templates,
                template_id=template_id,
                request_url=request_url,
                api_key=api_key,
                model_name=model_name,
                temperature=temperature,
                speed_mode=speed_mode,
                strict_reference_isolation=strict_reference_isolation,
                format_output=format_output,
            ))
        except Exception as exc:
            failures.append({
                "filename": material.get("name") or f"document_{index}",
                "reason": str(exc),
            })

    if not documents:
        detail = "；".join(f"{item['filename']}：{item['reason']}" for item in failures[:5]) or "未知原因"
        raise RuntimeError(f"全部材料生成失败。{detail}")

    first = documents[0]
    warnings = [warning for item in documents for warning in item.get("warnings", [])]
    if failures:
        warnings.append(
            f"部分材料未生成成功（{len(failures)}/{len(failures) + len(documents)}）："
            + "；".join(f"{item['filename']}：{item['reason']}" for item in failures[:3])
        )
    return {
        **first,
        "documents": documents,
        "failures": failures,
        "warnings": warnings,
        "timings": {"totalSeconds": round(perf_counter() - started_at, 2)},
        "batchMode": "one_file_one_document",
        "batchHint": "每个上传材料单独生成一份公文，不会自动汇总成一篇。",
    }


def _upload_meta_path(upload_path: Path) -> Path:
    return upload_path.with_suffix(upload_path.suffix + ".upload.json")


def _write_upload_meta(upload_path: Path, filename: str) -> None:
    meta_path = _upload_meta_path(upload_path)
    meta_path.write_text(
        json.dumps({"filename": Path(filename or "document").name}, ensure_ascii=False),
        encoding="utf-8",
    )


def encode_upload_filename(filename: str) -> str:
    """保留旧编码逻辑，供历史 temp 路径回退解析。"""
    safe_name = Path(filename or "document").name
    stem = Path(safe_name).stem or "document"
    suffix = Path(safe_name).suffix.lower()
    return f"{urllib.parse.quote(stem, safe='')}{suffix}"


def build_upload_temp_path(directory: Path, filename: str, suffix: str | None = None) -> Path:
    """落盘使用短 ASCII 路径，原始文件名写入 sidecar，避免 Windows 长路径与编码问题。"""
    safe_name = Path(filename or "document").name
    resolved_suffix = (suffix or Path(safe_name).suffix or ".bin").lower()
    token = uuid.uuid4().hex
    upload_path = directory / f"{token}{resolved_suffix}"
    _write_upload_meta(upload_path, safe_name)
    return upload_path


def display_name_from_temp_path(path: Path) -> str:
    meta_path = _upload_meta_path(path)
    if meta_path.exists():
        try:
            payload = json.loads(meta_path.read_text(encoding="utf-8"))
            original = str(payload.get("filename") or "").strip()
            if original:
                return original
        except (OSError, json.JSONDecodeError, TypeError):
            pass
    name = path.name
    if "_" not in name:
        return name
    _, encoded_name = name.split("_", 1)
    stem = Path(encoded_name).stem
    suffix = Path(encoded_name).suffix
    return f"{urllib.parse.unquote(stem)}{suffix}"
