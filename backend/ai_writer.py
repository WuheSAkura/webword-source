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
import shutil
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
# 短材料成文参考下限；实际篇幅按原材料长度比例估算，不再强制凑字。
MIN_DRAFT_CHARS = int(os.getenv("AI_MIN_DRAFT_CHARS", "800"))
# 成文硬上限（每份）
MAX_DRAFT_CHARS = int(os.getenv("AI_MAX_DRAFT_CHARS", "30000"))
# 无材料时的输出上限兜底；有材料时以原材料长度为基准。
MIN_OUTPUT_LIMIT_CHARS = int(os.getenv("AI_MIN_OUTPUT_LIMIT_CHARS", "2000"))
# 单次模型 HTTP 调用超时（秒）。高并发排队时单次推理可能很久，默认放宽到 10 分钟。
MODEL_HTTP_TIMEOUT_SECONDS = int(os.getenv("AI_MODEL_HTTP_TIMEOUT_SECONDS", "600"))
# 按速度档配置输出 token 预算：一次成文用足预算，避免多档递增重打模型。
DRAFT_TOKEN_BUDGETS = {
    "fast": (8192,),
    "standard": (12288,),
    "deep": (16384,),
}
# 完整材料直送模型；超过此上限时拒绝请求，绝不静默截断或自动摘要。
MAX_DIRECT_INPUT_CHARS = int(os.getenv("AI_MAX_DIRECT_INPUT_CHARS", "100000"))
MAX_REFERENCE_INPUT_CHARS = int(os.getenv("AI_MAX_REFERENCE_INPUT_CHARS", "50000"))
# 成文护栏：解析失败可重试；照录过高时单独计数再创作（默认最多 3 次），用尽后失败不放行。
MAX_GENERATION_GUARD_ATTEMPTS = int(os.getenv("AI_MAX_GENERATION_GUARD_ATTEMPTS", "3"))
MAX_COPY_REWRITE_ATTEMPTS = int(os.getenv("AI_MAX_COPY_REWRITE_ATTEMPTS", "3"))
SIMILARITY_REWRITE_COPY_THRESHOLD = float(os.getenv("AI_SIMILARITY_COPY_THRESHOLD", "0.90"))
# 段落近照录比例达到该阈值时，触发「禁止照录」再创作。
SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD = float(os.getenv("AI_SIMILARITY_PARAGRAPH_THRESHOLD", "0.35"))
# 合并链路：结构策划并入成文一次调用（默认开启）；材料整理默认单独调轻量模型。
AI_MERGE_PIPELINE = os.getenv("AI_MERGE_PIPELINE", "1").strip().lower() in {"1", "true", "yes", "on"}
# 兼容旧环境变量：显式 skip 仅在未开合并时生效。
AI_SKIP_PREP_LLM = os.getenv("AI_SKIP_PREP_LLM", "0").strip().lower() in {"1", "true", "yes", "on"}
AI_SKIP_PLAN_LLM = os.getenv("AI_SKIP_PLAN_LLM", "0").strip().lower() in {"1", "true", "yes", "on"}
AI_SKIP_FILL_LLM = os.getenv("AI_SKIP_FILL_LLM", "1").strip().lower() in {"1", "true", "yes", "on"}
# 成文提示里范文原文上限，避免超长上下文拖慢推理。
AI_REFERENCE_SAMPLE_CHARS = int(os.getenv("AI_REFERENCE_SAMPLE_CHARS", "2000"))
# 材料照录过高时是否再创作一轮（合并链路默认开启，保证不是原文粘贴）。
AI_REWRITE_ON_MATERIAL_SIMILARITY = os.getenv(
    "AI_REWRITE_ON_MATERIAL_SIMILARITY", "1"
).strip().lower() in {"1", "true", "yes", "on"}
# 成文默认不把原材料全文塞进提示（事实清单已含要点）；需要核对时可设 1。
AI_INCLUDE_RAW_MATERIAL_IN_DRAFT = os.getenv(
    "AI_INCLUDE_RAW_MATERIAL_IN_DRAFT", "0"
).strip().lower() in {"1", "true", "yes", "on"}
# 规则事实清单单条上限，避免「事实清单=原文逐段复制」。
AI_MATERIAL_FACT_SNIPPET_CHARS = int(os.getenv("AI_MATERIAL_FACT_SNIPPET_CHARS", "240"))
# 成稿正文与原材料正文的整体相似度（不含版头版脚）达到该值视为照录。
SIMILARITY_BODY_SOURCE_THRESHOLD = float(os.getenv("AI_SIMILARITY_BODY_SOURCE_THRESHOLD", "0.72"))
SIMILARITY_DRAFT_SOURCE_THRESHOLD = float(os.getenv("AI_SIMILARITY_DRAFT_SOURCE_THRESHOLD", "0.82"))
BODY_DRAFT_ROLES = frozenset({"body", "h1", "h2", "h3", "h4"})
HEADER_PRESERVE_ROLES = frozenset({
    "title", "subtitle", "recipient", "attachment_head",
    "sign_unit", "sign_date", "sign_contact", "security",
})
SIMILARITY_FAILURE_MARKERS = (
    "参考范文事实",
    "重合度过高",
    "重合度偏高",
    "材料原文重合",
    "禁止照录",
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
# 仅拦截明确的「待填」占位；不含「某某/某单位」——判决书脱敏人名、示例单位名常见于原材料。
RE_PLACEHOLDER_TEXT = re.compile(
    r"(请补充|待补|待定|待填写|请填写|"
    r"(?<![A-Za-z0-9])XXXX+(?![A-Za-z0-9])|"
    r"[\[【](?:[^\]】]{0,20})(?:主要领导|分管领导|处室名称|任务|目标|难点|方面|数量|时间)(?:[^\]】]{0,20})[\]】])",
    re.IGNORECASE,
)
# 整行几乎全是占位符时才在清洗阶段删除；「张某某」等材料原字予以保留。
RE_STANDALONE_PLACEHOLDER_LINE = re.compile(
    r"^(?:请补充|待补|待定|待填写|请填写|XXXX+|×{2,}|某单位|某某)+$",
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
RE_REASONING_META_DRAFT = re.compile(
    r"(?i)\b(need to analyze|let me analyze|the user gives|i should|first,? i'll|we need to)"
)
RE_EMPTY_MATERIAL_VALUE = re.compile(r"^(?:无|暂无|N/?A|-+)$", re.IGNORECASE)
RE_MATERIAL_METADATA = re.compile(
    r"(?:签发人|等级|发电时间|承办单位|发布日期|浏览次数|案号|案由|文号|发文字号)\s*[：:]?"
    r"|^[^\s]{1,24}〔\s*\d{4}\s*〕[^\s]{0,24}号$"
)
RE_REFERENCE_FACT_TOKEN = re.compile(
    r"\d+(?:\.\d+)+|\d{4,}|[〇一二三四五六七八九十]{4}年[〇一二三四五六七八九十月日]{1,8}"
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
    return build_plain_text_inventory(text)


def build_plain_text_inventory(text: str) -> dict[str, Any]:
    """由对话框粘贴文本直接构建材料清单，跳过文件解析。"""
    items: list[dict[str, Any]] = []
    stats = _new_stats()
    normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    for index, line in enumerate(normalized.splitlines(), start=1):
        if line.strip():
            _add_inventory_item(items, stats, "text", f"line {index}", line)
    if not items and normalized.strip():
        # 整段无换行时仍作为一条正文材料
        _add_inventory_item(items, stats, "text", "line 1", normalized.strip())
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


def inventory_plain_text(inventory: dict[str, Any]) -> str:
    """供模型阅读的纯正文，不含 [id|type|loc] 前缀，避免把清单标记当成正文。"""
    lines = []
    for item in inventory.get("items", []):
        text = str(item.get("text") or "").strip()
        if text and not RE_EMPTY_MATERIAL_VALUE.match(text):
            lines.append(text)
    return "\n".join(lines).strip()


def _text_overlap_ratio(left: str, right: str) -> float:
    """粗算两段文本重合度，用于识别「整理底稿≈原文照录」。"""
    left_shingles = _normalized_shingles(left or "", size=8)
    right_shingles = _normalized_shingles(right or "", size=8)
    if not left_shingles or not right_shingles:
        if not left or not right:
            return 0.0
        return SequenceMatcher(None, left[:4000], right[:4000]).ratio()
    return len(left_shingles & right_shingles) / max(len(left_shingles), 1)


DEFAULT_WRITING_PROMPT = (
    "请根据下列材料创作正式公文：以材料中的事实、数据、专名、日期与结论为依据，"
    "按目标文种设计篇章结构，用规范公文语体独立成文；"
    "须改写表述、重组段落，不得照搬或照录原材料底稿；"
    "主标题可据事由润写，不必沿用材料原标题原字；"
    "不编造材料未提供的数据、单位、日期及情节。"
)
# 兼容旧引用
DEFAULT_DRAFT_IMPROVE_PROMPT = DEFAULT_WRITING_PROMPT


def split_dialog_prompt_and_material(text: str) -> tuple[str, str]:
    """
    从对话框解析用户定制要求与材料正文。
    支持：
    - 【写作要求】…【材料】…
    - 短要求 + 单独一行 --- + 材料正文
    无标记时返回 ('', 全文)。
    """
    raw = (text or "").strip()
    if not raw:
        return "", ""
    req_match = re.search(r"【(?:写作)?要求】\s*", raw)
    mat_match = re.search(r"【材料】\s*", raw)
    if req_match and mat_match and req_match.start() < mat_match.start():
        user_prompt = raw[req_match.end():mat_match.start()].strip()
        material = raw[mat_match.end():].strip()
        if material:
            return user_prompt, material
    if mat_match and (not req_match or mat_match.start() < req_match.start()):
        material = raw[mat_match.end():].strip()
        user_prompt = raw[:mat_match.start()].strip()
        user_prompt = re.sub(r"^【(?:写作)?要求】\s*", "", user_prompt).strip()
        if material:
            return user_prompt, material
    sep = re.search(r"\n\s*-{3,}\s*\n", raw)
    if sep:
        left = raw[:sep.start()].strip()
        right = raw[sep.end():].strip()
        if right and left and len(left) <= 800:
            return left, right
    return "", raw


def format_user_directive(prompt: str, *, is_default: bool = False) -> str:
    """把用户提示包装为最高优先级指令块，供策划/成文提示使用。"""
    text = (prompt or "").strip() or DEFAULT_WRITING_PROMPT
    if is_default:
        return f"【默认创作指引】\n{text}"
    return (
        "【用户定制要求·最高优先级】\n"
        f"{text}\n"
        "须优先落实上述定制（侧重点、语气、篇幅、结构偏好、增删意向等）；"
        "与材料冲突时以用户要求的表达与取舍为准，但不得编造材料未提供的事实。"
    )


def resolve_writing_prompt(prompt: str, material_texts: list[str]) -> str:
    """解析写作要求：用户定制提示始终保留；空提示时使用默认创作指引。"""
    prompt_text = (prompt or "").strip()
    if not prompt_text:
        return DEFAULT_WRITING_PROMPT
    return prompt_text


def is_default_writing_prompt(prompt: str) -> bool:
    text = (prompt or "").strip()
    return (not text) or text == DEFAULT_WRITING_PROMPT

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
    path: Path | None,
    prompt: str,
    template_label: str,
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str,
    inventory: dict[str, Any] | None = None,
    name: str | None = None,
    allow_degradation: bool = False,
) -> tuple[dict[str, Any], list[str]]:
    """整理材料：文件路径走抽取；纯文本材料直接用 inventory，后续链路相同。"""
    if path is not None:
        display_name = name or display_name_from_temp_path(path)
        inventory = inventory or extract_document_inventory(path)
    else:
        display_name = name or "对话框材料.txt"
        if inventory is None:
            raise RuntimeError("纯文本材料缺少内容清单")
    raw_text = inventory_to_text(inventory)
    material_zones = analyze_material_zones(inventory, path, template_id)
    warnings: list[str] = list(inventory.get("warnings") or [])
    if path is None:
        warnings.append("材料来自对话框文本，已跳过文件提取，直接进入整理与成文")
    material_quality = assess_material_content(inventory)
    if material_quality["status"] != "ready":
        raise RuntimeError(f"材料正文不足：{material_quality['reason']}。请上传正文页、附件或完整文件，或在对话框粘贴更完整的材料")
    if len(raw_text) > MAX_DIRECT_INPUT_CHARS:
        raise RuntimeError(
            f"{display_name} 共 {len(raw_text)} 个字符，超过完整文本直送上限 {MAX_DIRECT_INPUT_CHARS}。"
            "为避免截断或自动摘要，已停止生成；请拆分材料或提高 AI_MAX_DIRECT_INPUT_CHARS。"
        )
    processed = process_upload_material_with_ai(
        inventory, prompt, template_label, template_id,
        request_url, api_key, model_name,
        material_zones=material_zones,
        allow_degradation=allow_degradation,
    )
    if processed["processingMode"] == "ai_structured_brief":
        prep_secs = processed.get("prepSeconds")
        time_note = f"，整理 {prep_secs}s" if prep_secs is not None else ""
        warnings.append(f"已完成材料事实抽取与要点整理{time_note}，后续依体例独立创作成文")
    else:
        prep_secs = processed.get("prepSeconds")
        time_note = f"（{prep_secs}s）" if prep_secs is not None else ""
        if AI_SKIP_PREP_LLM:
            warnings.append(f"材料整理走平台本地快路径{time_note}，已生成规则化事实清单")
        else:
            warnings.append(
                f"材料 AI 整理不可用{time_note}，已回退为规则化事实清单（不含原材料全文）"
            )
    prep_warning = str(processed.get("prepWarning") or "").strip()
    if prep_warning:
        warnings.append(prep_warning)
    if path is not None:
        source_elements = extract_source_elements(path, template_id, inventory)
    else:
        source_elements = extract_source_elements(Path("dialog_material.txt"), template_id, inventory)
    brief = processed["briefForDrafting"]
    material_digest = brief if brief and brief.strip() != raw_text.strip() else ""
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
        "name": display_name,
        "text": raw_text,
        "materialDigest": material_digest,
        "rawText": processed.get("rawText") or raw_text,
        "summary": processed.get("summary") or "",
        "title": prepared_title,
        "structuredFacts": processed.get("facts") or [],
        "attachmentBrief": processed.get("attachmentBrief") or "",
        "fullTextLength": len(raw_text),
        "rawTextLength": len(raw_text),
        "chunkCount": 1 if raw_text else 0,
        "processingMode": processed["processingMode"],
        "prepSeconds": processed.get("prepSeconds"),
        "prepAttempts": processed.get("prepAttempts"),
        "readReport": build_read_report(inventory),
        "materialQuality": material_quality,
        "materialZones": material_zones,
        "sourceElements": source_elements,
    }, warnings


def draft_char_limit(materials: list[dict[str, Any]], reference: dict[str, Any]) -> int:
    """按材料与体例估算成文上限；篇幅与原材料长度成比例，不强制虚高下限。"""
    source_chars = 0
    for item in materials:
        raw_chars = int(item.get("rawTextLength") or len(item.get("rawText") or item.get("text") or ""))
        source_chars += raw_chars
    reference_chars = int(reference.get("textChars") or len(reference.get("text", "")))
    if source_chars <= 0:
        proportional = MIN_OUTPUT_LIMIT_CHARS
    else:
        proportional = int(source_chars * 1.35)
    return min(
        MAX_DRAFT_CHARS,
        max(proportional, int(reference_chars * 1.15)),
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


def render_material_zones_text(
    zones: dict[str, Any],
    *,
    for_drafting: bool = False,
) -> str:
    """生成正向材料分区说明，帮助模型理解主件与附件边界。"""
    lines = [
        "上传材料结构识别（成文时按此分区取材）：",
        f"- 主件正文区（写入函/通知主体）：{len(zones.get('letterFactIds') or [])} 段事实",
        "- 成文完整性：独立主标题 [[title]]、主体正文、材料具备时的落款；"
        "「一、二、三…」只作 [[h1]]，不作主标题。",
    ]
    if zones.get("titleText") and not looks_like_section_heading(zones["titleText"]):
        if for_drafting:
            lines.append(
                "- 主件标题：材料含完整标题线索，成文时请据事实独立润写 [[title]]，"
                "概括事由即可，禁止照录原标题全文"
            )
        else:
            lines.append(f"- 主件标题（独立 [[title]] 段）：{zones['titleText']}")
    elif not zones.get("titleText") or looks_like_section_heading(zones.get("titleText") or ""):
        lines.append("- 主件标题：材料中需从会议/主题名称提炼独立 [[title]]（勿用一级小标题充当）")
    if zones.get("recipientText"):
        if for_drafting:
            lines.append("- 主送机关：材料含主送信息，成文时写入 [[recipient]]（专名按事实保留）")
        else:
            lines.append(f"- 主送机关（独立 [[recipient]] 段）：{zones['recipientText']}")
    if zones.get("attachmentHeadText"):
        if for_drafting:
            lines.append(
                "- 附件说明：材料含附件线索，成文时用一行 [[attachment_head]] 概括，"
                "禁止照录材料附件说明原句"
            )
        else:
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
    requested = str(requested_template_id or "").strip()
    if requested:
        return normalize_template_id(requested)
    combined = "\n".join([prompt, *file_texts])
    return normalize_template_id(guess_template_id(combined))


def build_template_key(template: dict[str, Any]) -> str:
    template_id = str(template.get("id") or "").strip()
    source_dir = str(template.get("sourceDir") or "").strip()
    if template_id and source_dir:
        return f"{template_id}::{source_dir}"
    return template_id


def parse_template_key(template_key: str | None) -> tuple[str, str]:
    raw = str(template_key or "").strip()
    if not raw:
        return "", ""
    if "::" in raw:
        template_id, source_dir = raw.split("::", 1)
        return template_id.strip(), source_dir.strip()
    return raw, ""


def get_template_by_key(
    template_key: str | None = None,
    template_id: str | None = None,
) -> dict[str, Any] | None:
    resolved_id, source_dir = parse_template_key(template_key)
    if not resolved_id:
        resolved_id = str(template_id or "").strip()
    if not resolved_id:
        return None
    templates = scan_template_library()
    if source_dir:
        exact = next(
            (item for item in templates if item.get("id") == resolved_id and item.get("sourceDir") == source_dir),
            None,
        )
        if exact:
            return exact
    matches = [item for item in templates if item.get("id") == resolved_id]
    if not matches:
        return None
    return max(
        matches,
        key=lambda item: (
            int((item.get("corpusProfile") or {}).get("availableReferenceCount") or 0),
            int(item.get("fileCount") or 0),
        ),
    )


def resolve_generation_template(
    prompt: str,
    file_texts: list[str],
    *,
    template_id: str | None = None,
    template_key: str | None = None,
) -> tuple[str, dict[str, Any]]:
    template = get_template_by_key(template_key, template_id)
    if template:
        return str(template["id"]), template
    resolved_id = choose_template_id(prompt, file_texts, template_id)
    template = get_template_by_key(None, resolved_id)
    if not template:
        raise RuntimeError(f"未找到目标公文模板：{resolved_id}")
    return resolved_id, template


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
        "title": "从材料会议名/事由名提炼整篇主标题，居中、无句号；不用「一、」类章节小标题",
        "subtitle": "材料有括注单位行时原字写入",
        "body": (
            "合并本段 sourceIds 同类事实，换用公文句式重组："
            "先背景/依据，再举措/成效；多项时用「一是…二是…」并列，压缩口语与重复过渡"
        ),
        "h1": "用「一、二、三…」概括本大段主题短语；后接独立 [[body]] 展开对应事实",
        "h2": "在大段内用「（一）（二）…」细分主题；序号随所属大段从（一）重计，后接 [[body]]",
        "h3": "概括细分要点主题短语",
        "h4": "概括细分要点主题短语",
        "recipient": "材料已有主送机关时顶格写出并加冒号",
        "attachment_head": "材料已有附件说明时原字一行写出",
        "sign_unit": "材料已有落款单位时逐行写出",
        "sign_date": "材料已有成文日期时用汉字日期写出",
        "sign_contact": "材料已有联系人及电话时原字写出",
    }
    return defaults.get(role, "按本段职责，合并对应事实后用公文语体重组表达")


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
            "[[title]]关于商请协助开展专项检查的函\n"
            "[[recipient]]各相关单位：\n"
            "[[body]]根据上级统一部署，我单位拟于近期开展专项检查。现就有关事项函告如下："
            "检查时间安排在3月中旬，重点核查台账资料与现场落实情况。\n"
            "[[body]]请各单位做好迎检准备，并于检查结束后5个工作日内反馈有关情况。\n"
            "[[attachment_head]]附件：检查清单\n"
            "[[sign_unit]]某市公安局\n"
            "[[sign_unit]]某分局办公室\n"
            "[[sign_date]]2024年3月1日\n"
            "[[sign_contact]]（联系人及电话：张三，010-12345678）\n"
        )
        heading_hint = "本篇文种以连续正文为主，版式重点在标题、主送、正文段与落款分区。"
        development = "正文一两段写清事由、事项与希望配合内容；分项可写在正文段内"
        rewrite_guide = (
            "函类创作：依据材料事实独立成文，开篇写清缘由与依据，主体连贯展开事项；"
            "多项并列写在正文段内「一是…二是…」；保留专名与数据，用公文语体组织论述，不编造未提供情节。"
        )
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
            "[[title]]关于报送2024年度工作总结材料的通知\n"
            "[[subtitle]]（办公室）\n"
            "[[h1]]一、总体工作情况\n"
            "[[body]]本年度围绕年度目标任务，扎实推进各项工作，取得阶段性成效。\n"
            "[[h1]]二、存在的主要问题\n"
            "[[body]]当前工作中仍存在基础台账不够规范、协同机制有待完善等问题。\n"
            "[[sign_unit]]某单位办公室\n"
            "[[sign_date]]2024年12月15日\n"
        )
        heading_hint = "可按材料复杂度设置一级/二级标题；标题措辞由材料归纳。"
        development = "按事项逻辑分块展开，标题统领正文"
        rewrite_guide = (
            "层级文种创作：依据材料事实独立成文，开篇用导语交代背景与目的；"
            "主体按 h1 分块，块内先总领再「一是…二是…」展开举措或成效；"
            "保留数字与专名，设计层次与衔接；收束可写成效概括或下步安排，不编造。"
        )
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
            f"[[title]]关于推进某项重点工作的{label}\n"
            "[[body]]为落实上级部署要求，现就有关事项明确如下。\n"
            "[[sign_unit]]某单位\n"
            "[[sign_date]]2024年6月1日\n"
        )
        heading_hint = "按材料复杂度决定是否分块；内容少时连续正文即可。"
        development = "按文种常规与材料事实组织主体"
        rewrite_guide = (
            "通用创作：依据材料事实独立成文，事项少则连贯成段，事项多则分块展开；"
            "保留硬事实（数字、日期、专名、结论），用公文语体组织论述，不编造。"
        )

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

    rewrite_example = (
        "材料要点：召开年度工作会议，部署三项重点任务。\n"
        "成文表达：一是召开年度工作会议，传达上级精神；二是围绕年度目标分解三项重点任务；"
        "三是明确责任分工与时限要求。"
    )

    return {
        "templateId": template_id,
        "label": label,
        "purpose": purpose,
        "mode": mode,
        "preferredRoles": preferred_roles,
        "chapterGuide": chapter_guide,
        "planGuide": plan_guide,
        "headingHint": heading_hint,
        "rewriteGuide": rewrite_guide,
        "rewriteExample": rewrite_example,
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

    candidate = text[start:end + 1]

    def _escape_literal_newlines_in_json_strings(s: str) -> str:
        """修复：LLM 往往在 JSON 字符串内部输出了“字面换行”，该 JSON 无法被 json.loads 解析。

        只在字符串内部把 '\n'/'\r' 变为 '\\n'，尽量不改变 JSON 结构。
        """
        out: list[str] = []
        in_string = False
        escape = False
        for ch in s:
            if in_string:
                if escape:
                    out.append(ch)
                    escape = False
                    continue
                if ch == "\\":
                    out.append(ch)
                    escape = True
                    continue
                if ch == '"':
                    in_string = False
                    out.append(ch)
                    continue
                if ch == "\n" or ch == "\r":
                    out.append("\\n")
                    continue
                out.append(ch)
            else:
                if ch == '"':
                    in_string = True
                out.append(ch)
        return "".join(out)

    def _try_load(s: str) -> dict[str, Any]:
        payload = json.loads(s)
        if not isinstance(payload, dict):
            raise RuntimeError("模型返回的 JSON 不是对象")
        return payload

    try:
        payload = _try_load(candidate)
    except json.JSONDecodeError:
        repaired = candidate
        # 容错：去掉尾逗号（JSON5 常见输出错误）
        repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
        # 容错：把字符串内部字面换行转为 \\n
        repaired = _escape_literal_newlines_in_json_strings(repaired)
        try:
            payload = _try_load(repaired)
        except Exception as exc:
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
    lines = ["【材料事实清单】（要点摘录，成文须改写表述，禁止照录）"]
    snippet_limit = max(AI_MATERIAL_FACT_SNIPPET_CHARS, 80)
    for item in inventory.get("items") or []:
        if item.get("sourceType") in {"header", "footer"}:
            continue
        text = str(item.get("text") or "").strip()
        if not text or RE_EMPTY_MATERIAL_VALUE.match(text):
            continue
        fact_id = str(item.get("id") or f"F{len(facts) + 1}")
        snippet = text
        if len(snippet) > snippet_limit:
            snippet = snippet[:snippet_limit].rstrip() + "…"
        facts.append({
            "id": fact_id,
            "category": str(item.get("sourceType") or "body"),
            "content": snippet,
            "location": str(item.get("location") or ""),
            "fullContent": text,
        })
        lines.append(f"[{fact_id} | {item.get('sourceType')}] {snippet}")
    brief = "\n".join(lines).strip()
    return brief, facts


_PREP_FIELD_ALIASES = {
    "标题": "title", "主标题": "title", "title": "title",
    "主送": "recipient", "主送机关": "recipient", "recipient": "recipient",
    "摘要": "summary", "summary": "summary",
    "附件": "attachmentBrief", "附件说明": "attachmentHead", "attachmentHead": "attachmentHead",
    "落款": "signUnits", "落款单位": "signUnits", "signUnits": "signUnits",
    "日期": "signDate", "成文日期": "signDate", "signDate": "signDate",
    "要点": "letterBrief", "综述": "letterBrief", "材料要点": "letterBrief",
    "letterBrief": "letterBrief", "briefForDrafting": "letterBrief",
}


def _split_tagged_prep_text(text: str) -> dict[str, str]:
    """按【标题】【要点】等中文标记切分自由文本，无标记时整篇作为要点。"""
    fields: dict[str, list[str]] = {}
    current = "letterBrief"
    for raw_line in text.replace("\r\n", "\n").split("\n"):
        line = raw_line.strip()
        match = re.match(r"^【(?P<label>[^】]{1,20})】\s*(?P<rest>.*)$", line)
        if match:
            label = match.group("label").strip()
            key = _PREP_FIELD_ALIASES.get(label)
            if key:
                current = key
                rest = match.group("rest").strip()
                if rest:
                    fields.setdefault(current, []).append(rest)
                else:
                    fields.setdefault(current, [])
                continue
        fields.setdefault(current, []).append(raw_line.rstrip())
    return {key: "\n".join(lines).strip() for key, lines in fields.items() if "\n".join(lines).strip()}


def _facts_from_prep_text(text: str) -> list[dict[str, Any]]:
    """从自由文本中回收事实条目，供平台继续处理。"""
    facts: list[dict[str, Any]] = []
    for match in re.finditer(
        r"(?:^|\n)\s*(?:[-*•]|\d+[、.．)]|\[(?P<fid>F?\d+)\])\s*(?P<body>[^\n]{6,400})",
        text,
    ):
        body = str(match.group("body") or "").strip(" ：:;-")
        if not body:
            continue
        fact_id = str(match.group("fid") or f"F{len(facts) + 1}").strip() or f"F{len(facts) + 1}"
        if not fact_id.startswith("F"):
            fact_id = f"F{fact_id}" if fact_id.isdigit() else f"F{len(facts) + 1}"
        facts.append({"id": fact_id, "zone": "letter", "category": "要点", "content": body})
    return facts


def _parse_material_prep_response(content: str) -> dict[str, Any]:
    """平台侧消化模型输出：自由文本优先，JSON 仅作可选补充。"""
    text = (content or "").strip()
    if not text:
        raise RuntimeError("材料整理模型没有返回内容")
    text = text.replace("\\n", "\n").strip()
    payload: dict[str, Any] = {}
    if "{" in text and "}" in text:
        try:
            payload = dict(_parse_json_object(text))
        except RuntimeError:
            payload = {}

    tagged = _split_tagged_prep_text(text)
    brief = str(
        tagged.get("letterBrief")
        or payload.get("letterBrief")
        or payload.get("briefForDrafting")
        or ""
    ).strip()
    if not brief:
        brief = text
    facts = payload.get("facts") if isinstance(payload.get("facts"), list) else []
    if not facts:
        facts = _facts_from_prep_text(brief) or _facts_from_prep_text(text)

    title = str(tagged.get("title") or payload.get("title") or "").strip()
    recipient = str(tagged.get("recipient") or payload.get("recipient") or "").strip()
    summary = str(tagged.get("summary") or payload.get("summary") or "").strip()
    attachment_head = str(tagged.get("attachmentHead") or payload.get("attachmentHead") or "").strip()
    attachment_brief = str(tagged.get("attachmentBrief") or payload.get("attachmentBrief") or attachment_head).strip()
    sign_date = str(tagged.get("signDate") or payload.get("signDate") or "").strip()
    sign_units = payload.get("signUnits") if isinstance(payload.get("signUnits"), list) else []
    tagged_units = str(tagged.get("signUnits") or "").strip()
    if tagged_units and not sign_units:
        sign_units = [part.strip() for part in re.split(r"[\n；;]+", tagged_units) if part.strip()]

    return {
        "summary": summary,
        "title": title,
        "recipient": recipient,
        "attachmentHead": attachment_head,
        "attachmentBrief": attachment_brief,
        "signUnits": sign_units,
        "signDate": sign_date,
        "facts": facts,
        "letterBrief": brief,
    }


def process_upload_material_with_ai(
    inventory: dict[str, Any],
    prompt: str,
    template_label: str,
    template_id: str,
    request_url: str,
    api_key: str,
    model_name: str,
    material_zones: dict[str, Any] | None = None,
    allow_degradation: bool = False,
) -> dict[str, Any]:
    """上传材料预处理：模型自由文本整理 + 平台解析；失败时用本地抽取继续成文。

    不再强制「短 JSON + <<<BRIEF>>>」分段。模型返回任意可读文本即可，由平台切分字段。
    """
    tagged_text = inventory_to_text(inventory)
    plain_text = inventory_plain_text(inventory) or tagged_text
    zones = material_zones or analyze_material_zones(inventory, None, template_id)
    zones_text = render_material_zones_text(zones)
    fallback_brief, fallback_facts = build_rule_based_material_brief(inventory)
    writing_prompt = resolve_writing_prompt(prompt, [plain_text])
    user_directive = format_user_directive(
        writing_prompt, is_default=is_default_writing_prompt(writing_prompt),
    )
    prep_started = perf_counter()
    material_input = plain_text

    def _local_pack(warning: str, attempt_count: int) -> dict[str, Any]:
        zone_title = str(zones.get("titleText") or "").strip()
        if looks_like_section_heading(zone_title):
            zone_title = ""
        return {
            "summary": zones.get("summary") or "",
            "facts": fallback_facts,
            "briefForDrafting": fallback_brief or str(zones.get("summary") or "").strip(),
            "attachmentBrief": "",
            "title": zone_title,
            "recipient": zones.get("recipientText") or "",
            "attachmentHead": zones.get("attachmentHeadText") or "",
            "signUnits": zones.get("signUnits") or [],
            "signDate": zones.get("signDateText") or "",
            "rawText": plain_text,
            "materialZones": zones,
            "processingMode": "rule_based_brief",
            "prepOverlap": None,
            "prepAttempts": attempt_count,
            "prepSeconds": round(perf_counter() - prep_started, 2),
            "prepWarning": warning,
        }

    # 显式跳过整理模型时，走本地规则抽取（降级路径）。
    if AI_SKIP_PREP_LLM:
        return _local_pack("材料整理走平台本地快路径（未调用整理模型）", attempt_count=0)

    def _pack_success(payload: dict[str, Any], attempt_count: int) -> dict[str, Any]:
        brief = str(payload.get("letterBrief") or "").strip() or fallback_brief
        facts = payload.get("facts") if isinstance(payload.get("facts"), list) else []
        if not facts:
            facts = fallback_facts
        summary = str(payload.get("summary") or zones.get("summary") or "").strip()
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
            "rawText": plain_text,
            "materialZones": zones,
            "processingMode": "ai_structured_brief",
            "prepOverlap": None,
            "prepWarning": "",
            "prepAttempts": attempt_count,
            "prepSeconds": round(perf_counter() - prep_started, 2),
        }

    def _run_prep(system_extra: str = "", temperature: float = 0.3) -> tuple[dict[str, Any], dict[str, Any]]:
        messages = [
            {
                "role": "system",
                "content": (
                    "你是机关公文材料分析师。任务是对上传材料做轻量整理，输出供公文创作使用的事实清单。\n"
                    "只输出纯文本，不要 JSON、不要 Markdown 代码块、不要 <<<BRIEF>>> 分隔符。\n"
                    "可用标记（有则写，无则省略）：【标题】【主送】【摘要】【落款】【日期】【附件】【要点】\n"
                    "【要点】用条目列出全部硬事实（主体、动作、数字、日期、专名、结论），"
                    "每条一句概括，禁止照录原段落或原句。\n"
                    "【摘要】一句话说明材料形态与取材策略（如零散素材/已成稿函件等），"
                    "由你判断后续成文应侧重提炼还是改写，无需用户选择模式。\n"
                    "【标题】仅标注材料主题/事由线索，供成文时润写主标题参考，不是成稿标题。\n"
                    "不编造、不评价、不扩写；用户定制要求优先。"
                    + (f"\n{system_extra}" if system_extra else "")
                ),
            },
            {
                "role": "user",
                "content": (
                    f"目标文种：{template_label}（{template_id}）\n"
                    f"{user_directive}\n\n"
                    f"{zones_text}\n\n"
                    f"原始材料（创作依据）：\n{material_input}\n\n"
                    "请直接输出整理结果。"
                ),
            },
        ]
        # 控制输出预算，避免思考链占满后正文为空
        prep_tokens = min(12288, max(4096, int(len(plain_text) * 0.8) + 2048))
        result = call_chat_model_result(
            request_url, api_key, model_name, messages,
            temperature=temperature, max_tokens=prep_tokens,
            thinking="disabled",
            reasoning_effort="none",
        )
        content = str(result.get("content") or "").strip()
        if not content:
            finish = result.get("finishReason") or "unknown"
            reasoning_chars = result.get("reasoningChars") or 0
            raise RuntimeError(
                f"材料整理模型没有返回可用正文（finish_reason={finish}，reasoning={reasoning_chars} 字）"
            )
        return _parse_material_prep_response(content), result

    last_error = ""
    last_meta: dict[str, Any] = {}
    try:
        payload, last_meta = _run_prep()
        return _pack_success(payload, attempt_count=1)
    except Exception as first_exc:
        last_error = str(first_exc)
        try:
            payload, last_meta = _run_prep(
                system_extra=(
                    "上一稿未返回可用正文。请仅输出纯文本材料整理结果；"
                    "务必包含【要点】段落；不要 JSON、不要代码块。"
                    "若材料含案件或敏感表述，仍只做信息摘录，不评价。"
                ),
                temperature=0.2,
            )
            return _pack_success(payload, attempt_count=2)
        except Exception as second_exc:
            last_error = f"{last_error}；重试：{second_exc}"

    # 模型空返回 / 网关风控 / 解析失败：用平台本地抽取继续，不拦死成文
    finish = last_meta.get("finishReason") or "unknown"
    warning = (
        f"材料 AI 整理未返回可用正文（finish_reason={finish}；{last_error[:180]}），"
        f"已改用平台本地事实抽取（规则降级，不含原材料全文）"
    )
    return _local_pack(warning, attempt_count=2)


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
        provisional_id = resolve_category_template_id(folder)
        references = [
            _ingest_reference(
                path, provisional_id,
                enrich_with_ai=enrich_structures,
                model_config=model_config,
            )
            for path in files
        ]
        template_id = resolve_category_template_id(folder, references)
        catalog_item = next((v for v in catalog_by_label.values() if v["id"] == template_id), None)
        for reference in references:
            reference["templateId"] = template_id
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


CATEGORY_META_FILENAME = ".category_meta.json"


def read_category_meta(folder: Path) -> dict[str, Any]:
    meta_path = folder / CATEGORY_META_FILENAME
    if not meta_path.is_file():
        return {}
    try:
        payload = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def write_category_meta(folder: Path, template_id: str, display_name: str = "") -> None:
    folder.mkdir(parents=True, exist_ok=True)
    meta = {
        "templateId": normalize_template_id(template_id),
        "displayName": (display_name or folder.name).strip(),
    }
    (folder / CATEGORY_META_FILENAME).write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def resolve_category_template_id(folder: Path, references: list[dict[str, Any]] | None = None) -> str:
    meta = read_category_meta(folder)
    meta_id = normalize_template_id(str(meta.get("templateId") or ""))
    if meta_id and meta_id != "generic":
        return meta_id
    folder_guess = normalize_template_id(guess_template_id("", folder.name))
    if folder_guess != "generic":
        return folder_guess
    for item in references or []:
        text = str(item.get("text") or "")
        filename = str(item.get("filename") or "")
        guessed = normalize_template_id(guess_template_id(text, filename))
        if guessed != "generic":
            return guessed
    return folder_guess


def sanitize_category_folder_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]', "", (name or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    if not cleaned:
        raise ValueError("模板分类名称不能为空")
    return cleaned[:80]


def resolve_category_template_id_for_upload(
    target_dir: Path,
    existing: dict[str, Any] | None = None,
    fallback_template_id: str = "",
) -> str:
    """上传归入已有分类时，以分类元数据/索引为准，不按文件内容改文种。"""
    meta_id = normalize_template_id(str(read_category_meta(target_dir).get("templateId") or ""))
    if meta_id:
        return meta_id
    if existing and existing.get("id"):
        return normalize_template_id(str(existing["id"]))
    if fallback_template_id:
        return normalize_template_id(fallback_template_id)
    folder_guess = normalize_template_id(guess_template_id("", target_dir.name))
    return folder_guess if folder_guess != "generic" else "generic"


def resolve_template_upload_target(
    *,
    template_id: str = "",
    source_dir: str = "",
    category_name: str = "",
    category_mode: str = "existing",
    probe_text: str = "",
    probe_filename: str = "",
) -> tuple[Path, str, str]:
    """解析模板上传目标目录、文种与展示名。"""
    mode = (category_mode or "existing").strip().lower()
    catalog = {item["id"]: item for item in get_template_catalog()}
    base_dir = TEMPLATE_SOURCE_DIR.resolve()

    if mode == "new":
        folder_name = sanitize_category_folder_name(category_name)
        resolved_id = normalize_template_id((template_id or "").strip())
        if not resolved_id or resolved_id == "generic":
            resolved_id = normalize_template_id(guess_template_id(probe_text, folder_name))
        if resolved_id == "generic":
            resolved_id = normalize_template_id(guess_template_id(probe_text, probe_filename))
        target_dir = (TEMPLATE_SOURCE_DIR / folder_name).resolve()
        if base_dir not in target_dir.parents and target_dir != base_dir:
            raise ValueError("非法模板目录")
        target_dir.mkdir(parents=True, exist_ok=True)
        write_category_meta(target_dir, resolved_id, folder_name)
        return target_dir, resolved_id, folder_name

    resolved_source = (source_dir or "").strip()
    if resolved_source:
        target_dir = Path(resolved_source).resolve()
        if base_dir not in target_dir.parents:
            raise ValueError("所选模板分类不存在")
        if not target_dir.is_dir():
            raise ValueError("所选模板分类不存在")
        existing = next(
            (item for item in scan_template_library() if Path(item.get("sourceDir") or "").resolve() == target_dir),
            None,
        )
        resolved_id = resolve_category_template_id_for_upload(
            target_dir,
            existing,
            fallback_template_id=(template_id or "").strip(),
        )
        display = str(existing.get("name") if existing else target_dir.name)
        return target_dir, resolved_id, display

    resolved_id = normalize_template_id((template_id or "").strip())
    if not resolved_id:
        resolved_id = normalize_template_id(guess_template_id(probe_text, probe_filename))
    if resolved_id == "generic":
        raise ValueError("自动判断无法识别文种，请选择模板分类或目标文种后上传")
    template = get_template_by_id(resolved_id)
    if not template:
        raise ValueError(f"未找到模板类别 {resolved_id}")
    target_dir = Path(template["sourceDir"]).resolve()
    display = str(template.get("name") or template.get("label") or resolved_id)
    return target_dir, resolved_id, display


def delete_template_file(source_dir: str, filename: str) -> None:
    target_dir = Path(source_dir).resolve()
    base_dir = TEMPLATE_SOURCE_DIR.resolve()
    if base_dir not in target_dir.parents:
        raise ValueError("非法模板目录")
    safe_name = Path(filename or "").name
    if not safe_name or safe_name.startswith("~$"):
        raise ValueError("无效文件名")
    file_path = (target_dir / safe_name).resolve()
    if base_dir not in file_path.parents or not file_path.is_file():
        raise ValueError("模板文件不存在")
    file_path.unlink()


def delete_template_category(source_dir: str) -> None:
    target_dir = Path(source_dir).resolve()
    base_dir = TEMPLATE_SOURCE_DIR.resolve()
    if base_dir not in target_dir.parents or target_dir == base_dir:
        raise ValueError("非法模板目录")
    if not target_dir.is_dir():
        raise ValueError("模板分类不存在")
    shutil.rmtree(target_dir)


def public_template_view(template: dict[str, Any]) -> dict[str, Any]:
    """模板正文仅供后端生成使用，模板接口只暴露索引元数据。"""
    hidden = {"samplePath", "sampleText"}
    view = {key: value for key, value in template.items() if key not in hidden}
    view["templateKey"] = build_template_key(template)
    return view


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


_SEMANTIC_STOP_TERMS = frozenset({
    "工作", "情况", "有关", "进行", "开展", "单位", "部门", "要求", "通知", "报告",
    "意见", "方案", "措施", "推进", "落实", "加强", "组织", "建设", "管理", "服务",
    "相关", "各项", "全面", "切实", "进一步", "不断", "坚持", "提高", "确保",
})


def _normalize_for_semantic(text: str) -> str:
    return "".join(re.findall(r"[\u4e00-\u9fffA-Za-z0-9]+", (text or "").lower()))


def _extract_semantic_terms(text: str, *, max_terms: int = 160) -> set[str]:
    terms: set[str] = set()
    for token in re.findall(r"[\u4e00-\u9fff]{2,8}", text or ""):
        if token not in _SEMANTIC_STOP_TERMS:
            terms.add(token)
    for token in re.findall(r"\d{4,}", text or ""):
        terms.add(token)
    if len(terms) > max_terms:
        return set(list(terms)[:max_terms])
    return terms


def _semantic_similarity_score(query: str, document: str) -> float:
    """字符序列 + 关键词覆盖；限制比对长度，避免范文匹配拖慢整链。"""
    if not query or not document:
        return 0.0
    q_norm = _normalize_for_semantic(query)[:4000]
    d_norm = _normalize_for_semantic(document)[:4000]
    if not q_norm or not d_norm:
        return 0.0
    seq_ratio = SequenceMatcher(None, q_norm, d_norm).ratio()
    q_terms = _extract_semantic_terms(query[:6000], max_terms=80)
    d_terms = _extract_semantic_terms(document[:6000], max_terms=80)
    if not q_terms:
        return round(seq_ratio, 4)
    union = q_terms | d_terms
    jaccard = len(q_terms & d_terms) / len(union) if union else 0.0
    coverage = len(q_terms & d_terms) / len(q_terms)
    return round(coverage * 0.5 + jaccard * 0.2 + seq_ratio * 0.3, 4)


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
            # 匹配只用前段，避免长材料拖慢语义打分
            material_bits.append(text[:4000])
    query = "\n".join([prompt, *material_bits]).strip() or prompt
    # 查询侧同样截断，保证匹配阶段亚秒级完成
    query = query[:6000]
    ngram_scores = _reference_similarity_scores(query, references)
    semantic_scores = [_semantic_similarity_score(query, item["text"]) for item in references]
    scored = []
    for index, reference in enumerate(references):
        multi_score, dimensions = _template_feature_score(query, prompt, reference, template["id"])
        final_score = round(
            semantic_scores[index] * 0.45 + multi_score * 0.35 + ngram_scores[index] * 0.20,
            4,
        )
        scored.append((final_score, {**dimensions, "semantic": semantic_scores[index], "ngram": ngram_scores[index]}))
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


def _facts_inventory_text(materials: list[dict[str, Any]], limit: int | None = None) -> str:
    """将结构化事实整理为成文完整性清单（默认不截断）。"""
    lines: list[str] = []
    for item in materials:
        for fact in item.get("structuredFacts") or []:
            fact_id = str(fact.get("id") or "").strip()
            content = str(fact.get("content") or "").strip()
            if not content:
                continue
            zone = str(fact.get("zone") or fact.get("category") or "").strip()
            prefix = f"[{fact_id}" + (f"|{zone}" if zone else "") + "]"
            lines.append(f"{prefix} {content}")
            if limit is not None and len(lines) >= limit:
                return "\n".join(lines)
    return "\n".join(lines)


def _material_context_text(
    materials: list[dict[str, Any]],
    *,
    include_raw: bool | None = None,
) -> str:
    """成文用材料上下文：事实要点优先；原文默认不进入提示，避免模型照录。"""
    use_raw = AI_INCLUDE_RAW_MATERIAL_IN_DRAFT if include_raw is None else bool(include_raw)
    parts = []
    for index, item in enumerate(materials, start=1):
        zones = item.get("materialZones") or {}
        zones_text = render_material_zones_text(zones, for_drafting=True) if zones else ""
        attachment_note = ""
        if item.get("attachmentBrief"):
            attachment_note = f"\n附件区摘要（不写入函件正文）：\n{item['attachmentBrief']}"
        mode = item.get("processingMode") or ""
        facts_text = _facts_inventory_text([item])
        raw_text = str(item.get("rawText") or item.get("text") or "").strip()
        digest = str(item.get("materialDigest") or item.get("summary") or "").strip()
        creation_guide = (
            "成文取用优先级：①用户定制要求（最高）→②硬事实清单→③材料要点。"
            "材料是事实素材，不是成稿底稿：须提炼主题、重组结构、用公文语体独立表述；"
            "禁止整段/整句照录原材料；硬事实（数字、专名、日期、结论）须保留且不得改写数值；"
            "不编造材料未提供的数据/单位/日期/情节。"
        )
        if use_raw:
            creation_guide += "文末附原材料核对节，仅供查漏，成文不得粘贴。"
        else:
            creation_guide += "当前未提供原材料全文，只能依据事实清单创作并改写表述。"
        if mode == "rule_based_brief":
            creation_guide += "事实清单为规则摘录，须改写句式后成文。"
        elif mode == "ai_structured_brief":
            creation_guide += "整理模型已输出抽象事实要点，请按文种体例独立成文，禁止照录底稿。"
        block_parts = [
            f"[Material {index}: {item['name']}]",
            creation_guide,
            zones_text,
        ]
        if facts_text:
            block_parts.append(f"硬事实清单（成文主依据；须覆盖相关事实并改写表述）：\n{facts_text}")
        if digest and digest != raw_text:
            block_parts.append(f"材料要点综述（辅助取材，须改写表述）：\n{digest}")
        if use_raw and raw_text:
            raw_cap = min(len(raw_text), 12000)
            raw_excerpt = raw_text[:raw_cap]
            if len(raw_text) > raw_cap:
                raw_excerpt += "\n…（原材料已截断，其余事实以清单为准）"
            block_parts.append(
                "原材料核对节（禁止粘贴照录，仅用于核对数字/专名/日期是否覆盖）：\n"
                f"{raw_excerpt}"
            )
        elif raw_text and not facts_text:
            block_parts.append(
                f"材料规模约 {len(raw_text)} 字；请据用户要求与文种体例独立成文，勿还原原文段落。"
            )
        if attachment_note:
            block_parts.append(attachment_note.strip())
        parts.append("\n".join(part for part in block_parts if part).strip())
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
    rewrite_guide = genre_profile.get("rewriteGuide") or ""
    user_directive = format_user_directive(
        prompt, is_default=is_default_writing_prompt(prompt),
    )
    system = (
        "你是公文结构策划助手。基于材料事实，给出可执行的正式公文段落蓝图。\n"
        f"当前文种「{genre_profile.get('label') or template_id}」：{genre_profile.get('purpose', '')}。"
        f"{genre_profile.get('planGuide', '')}"
        f"{rewrite_guide}"
        "用户定制要求优先级最高。蓝图必须含独立 title；「一、二、三…」用 h1；材料有落款则写 sign_unit/sign_date。\n"
        "只输出纯文本段落，每行一个角色，格式：[[role]]本段写法。不要 JSON，不要 Markdown。"
    )
    reference_sample = str(reference.get("text") or "").strip()
    user = (
        f"目标文种：{genre_profile.get('label') or template.get('label') or template.get('name')}（{template_id}）\n"
        f"文种用途：{genre_profile.get('purpose', '')}\n"
        f"写作品格：{genre_profile.get('mode', 'adaptive')}；常用角色：{preferred}\n"
        f"组织方式：{json.dumps(genre_profile.get('organization') or {}, ensure_ascii=False)}\n"
        f"创作指引：{rewrite_guide}\n"
        f"{user_directive}\n"
        f"主件正文区可用来源编号：{allowed_ids}\n\n"
        f"{zones_notes}\n\n"
        "模板结构化框架（JSON，文种体例坐标）：\n<template_structure>\n"
        f"{reference_outline}\n</template_structure>\n\n"
        + (f"体例样本原文（参照篇章骨架与写法）：\n<reference_sample>\n{reference_sample}\n</reference_sample>\n\n" if reference_sample else "")
        + "上传材料（事实清单优先；原文仅供核对，禁止照录）：\n<source>\n"
        f"{material_text or prompt}\n</source>\n\n"
        f"请按常用角色输出蓝图，例如：\n"
        "[[title]]据材料事由独立润写主标题\n"
        "[[recipient]]材料有主送则写入（专名按事实保留）\n"
        "[[body]]开篇写清缘由与依据\n"
        "[[h1]]按材料大段主题分块\n"
        "[[body]]覆盖对应硬事实并组织论述\n"
        "[[sign_unit]]材料有落款则写入（专名按事实保留）\n"
        "[[sign_date]]材料有日期则写入（按事实保留）"
    )
    return [{"role": "system", "content": system}, {"role": "user", "content": user}]


def _normalize_plan_sections(
    sections: list[dict[str, Any]], materials: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    valid_ids = set(_source_ids(materials))
    letter_ids = _letter_source_ids(materials) or list(valid_ids)
    normalized: list[dict[str, Any]] = []
    for index, section in enumerate(sections, start=1):
        if not isinstance(section, dict):
            continue
        role = str(section.get("role") or "")
        if role not in ALLOWED_DRAFT_ROLES:
            continue
        source_ids = [str(item) for item in (section.get("sourceIds") or []) if str(item) in valid_ids]
        if not source_ids:
            source_ids = letter_ids if role in {"body", "h1", "h2", "h3", "h4"} else list(valid_ids)[:3]
        normalized.append({
            "order": int(section.get("order") or index),
            "role": role,
            "function": str(section.get("function") or _default_section_function(role)),
            "writingMethod": str(section.get("writingMethod") or _default_writing_method(role)),
            "sourceIds": source_ids,
        })
    return normalized


def _plan_from_role_lines(content: str, materials: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """把 [[role]]写法 自由文本转成段落蓝图。"""
    sections: list[dict[str, Any]] = []
    for raw_line in content.replace("\r\n", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        pair = _match_role_line(line)
        role = ""
        text = ""
        if pair:
            role, text = pair
        else:
            marker = RE_ROLE_MARKER.match(line) or RE_ALT_ROLE_MARKER.match(line)
            if marker:
                role = _normalize_role_token(marker.group("role")) or ""
        if not role:
            continue
        sections.append({
            "order": len(sections) + 1,
            "role": role,
            "function": _default_section_function(role),
            "writingMethod": text or _default_writing_method(role),
            "sourceIds": [],
        })
    return _normalize_plan_sections(sections, materials)


def parse_transformation_plan(content: str, materials: list[dict[str, Any]]) -> dict[str, Any]:
    text = (content or "").strip()
    if not text:
        raise RuntimeError("范文结构分析没有返回内容")
    if "```" in text:
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.IGNORECASE | re.DOTALL).strip()

    plan: dict[str, Any] = {}
    start, end = text.find("{"), text.rfind("}")
    if start >= 0 and end > start:
        try:
            plan = json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            try:
                plan = _parse_json_object(text)
            except RuntimeError:
                plan = {}

    sections = plan.get("targetSections") if isinstance(plan.get("targetSections"), list) else []
    normalized_sections = _normalize_plan_sections(sections, materials) if sections else []
    if not normalized_sections:
        normalized_sections = _plan_from_role_lines(text, materials)
    if not normalized_sections:
        raise RuntimeError("范文结构分析未返回可用段落蓝图")

    organization = plan.get("organization") if isinstance(plan.get("organization"), dict) else {}
    return {
        "documentPurpose": str(plan.get("documentPurpose") or "按目标文种形成正式公文"),
        "organization": {
            "opening": str(organization.get("opening") or "按文种开头职责组织材料事实"),
            "development": str(organization.get("development") or "按文种主体职责展开材料事实"),
            "closing": str(organization.get("closing") or "按文种结尾职责收束材料事实"),
            "style": [str(item) for item in (organization.get("style") or []) if str(item)],
        },
        "targetSections": normalized_sections,
    }


def _roles_from_reference_structure(reference: dict[str, Any] | None) -> list[str]:
    """从范文结构化缓存提取段落角色序列。"""
    if not reference:
        return []
    blueprint = reference.get("structureJson") or {}
    roles = [str(role) for role in (blueprint.get("roleSequence") or []) if str(role)]
    if not roles:
        roles = [
            str(section.get("role"))
            for section in (blueprint.get("sections") or [])
            if section.get("role")
        ]
    compact: list[str] = []
    for role in roles:
        if role in ALLOWED_DRAFT_ROLES and (not compact or compact[-1] != role or role in {"body", "sign_unit"}):
            compact.append(role)
    return compact


def _material_h1_headings(materials: list[dict[str, Any]]) -> list[str]:
    """从原材料中识别「一、二、三…」一级标题。"""
    headings: list[str] = []
    for item in materials:
        text = str(item.get("rawText") or item.get("text") or "")
        for line in text.splitlines():
            line = line.strip()
            match = RE_H1_TEXT.match(line)
            if match:
                heading = match.group("text").strip()
                if heading and heading not in headings:
                    headings.append(heading)
    return headings


def build_fallback_transformation_plan(
    template_id: str,
    materials: list[dict[str, Any]],
    reference: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """结构策划 JSON 解析失败时，按范文体例 + 文种 + 材料分节生成完整段落蓝图。"""
    genre_profile = build_genre_writing_profile(template_id)
    ref_roles = _roles_from_reference_structure(reference)
    preferred_roles = ref_roles or list(genre_profile.get("preferredRoles") or ["title", "body", "sign_unit", "sign_date"])
    if "title" not in preferred_roles:
        preferred_roles.insert(0, "title")
    material_h1s = _material_h1_headings(materials)
    if material_h1s and "h1" in set(preferred_roles):
        expanded: list[str] = []
        for role in preferred_roles:
            if role == "body" and expanded and expanded[-1] not in {"h1", "h2", "title", "recipient", "subtitle"}:
                for _ in material_h1s:
                    expanded.extend(["h1", "body"])
            else:
                expanded.append(role)
        preferred_roles = expanded
    source_ids = _source_ids(materials)
    body_source_ids = _letter_source_ids(materials) or source_ids
    header_roles = {"title", "subtitle", "recipient", "attachment_head", "sign_unit", "sign_date", "sign_contact", "security"}
    sections = []
    h1_index = 0
    for order, role in enumerate(preferred_roles, start=1):
        if role in {"body", "h1", "h2", "h3", "h4"}:
            sid = body_source_ids
        elif role in header_roles:
            sid = source_ids[: max(1, min(3, len(source_ids)))]
        else:
            sid = source_ids
        writing_method = _default_writing_method(role)
        if role == "h1" and h1_index < len(material_h1s):
            writing_method = f"以「{material_h1s[h1_index]}」为大段主题，组织对应材料事实"
            h1_index += 1
        sections.append({
            "order": order,
            "role": role,
            "function": _default_section_function(role),
            "writingMethod": writing_method,
            "sourceIds": sid,
        })
    organization = genre_profile.get("organization") or {}
    if reference and (reference.get("structureJson") or {}).get("organization"):
        organization = (reference.get("structureJson") or {}).get("organization") or organization
    return {
        "documentPurpose": genre_profile.get("purpose") or "按目标文种形成正式公文",
        "organization": {
            "opening": str(organization.get("opening") or "按文种开头职责组织材料事实"),
            "development": str(organization.get("development") or "按文种主体职责展开材料事实"),
            "closing": str(organization.get("closing") or "按文种结尾职责收束材料事实"),
            "style": [str(item) for item in (organization.get("style") or []) if str(item)],
        },
        "targetSections": sections,
    }


def build_ai_messages(
    prompt: str, materials: list[dict[str, Any]], template: dict | None,
    reference: dict[str, Any], transformation_plan: dict[str, Any],
    *,
    omit_reference_body: bool = False,
    rewrite_mode: bool = False,
    rewrite_reason: str = "",
    merge_pipeline: bool | None = None,
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
        f"版式设计：{heading_hint}\n"
        f"创作指引：{genre_profile.get('rewriteGuide') or ''}"
    )

    material_text = _material_context_text(
        materials,
        include_raw=False if rewrite_mode else AI_INCLUDE_RAW_MATERIAL_IN_DRAFT,
    )
    zones_notes = "\n\n".join(
        render_material_zones_text(item.get("materialZones") or {}, for_drafting=True)
        for item in materials
        if item.get("materialZones")
    )
    use_merge = AI_MERGE_PIPELINE if merge_pipeline is None else bool(merge_pipeline)
    rewrite_note = ""
    if rewrite_mode:
        rewrite_note = (
            "【再生成·禁止照录】上一稿与材料段落近照录或质量不足。"
            f"触发原因：{rewrite_reason or '表述质量或段落照录风险'}。"
            "必须整篇重写句式与衔接：保留硬事实，换表达、换结构层次；"
            "正文不得与原材料大段相同；不编造材料未提供的事实。"
        )
    chapter_guide = genre_profile.get("chapterGuide") or (
        "按材料主题决定连续正文或分块标题；标题后紧跟正文。"
    )
    creation_guide = genre_profile.get("rewriteGuide") or (
        "依据材料事实独立创作：保留硬事实，设计篇章结构，用公文语体组织论述，不编造。"
    )
    creation_example = genre_profile.get("rewriteExample") or (
        "材料要点：召开年度工作会议，部署三项重点任务。\n"
        "成文表达：一是召开年度工作会议，传达上级精神；二是围绕年度目标分解三项重点任务；"
        "三是明确责任分工与时限要求。"
    )
    example = genre_profile.get("example") or (
        "[[title]]关于推进某项重点工作的公文\n[[body]]为落实上级部署要求，现就有关事项明确如下。\n"
    )
    source_chars = sum(
        int(item.get("rawTextLength") or len(item.get("rawText") or item.get("text") or ""))
        for item in materials
    )
    # 成文篇幅：材料转公文应压缩重组，不宜与原文等长（此前 1.15 倍会诱导照录）。
    creation_target = (
        min(output_limit, max(600, int(source_chars * 0.55)))
        if source_chars else output_limit
    )
    user_directive = format_user_directive(
        prompt, is_default=is_default_writing_prompt(prompt),
    )
    merge_workflow = ""
    if use_merge:
        merge_workflow = (
            "【合并工作流·一次完成】材料事实已由整理阶段抽取为硬事实清单。"
            "在输出成稿前，在推理中依次完成："
            "①结构策划：按目标文种与下方蓝图确定段落角色与写法；"
            "②独立成文：以事实清单为依据，用公文语体写出全文。"
            "只输出最终结构化段落，不要输出整理笔记或策划说明。\n\n"
        )
    anti_copy = (
        "【严禁照录】\n"
        "材料是事实素材不是底稿。禁止整段、整句复制或轻微改词粘贴原材料；"
        "须改写句式、重组段落与逻辑衔接；硬事实保留，叙述必须原创。"
        "若输出与材料大段相同，视为不合格。\n\n"
    )
    system = (
        "你是机关公文撰稿助手。基于上传材料事实，按目标文种创作可直接使用的正式公文。\n\n"
        f"{merge_workflow}"
        f"{anti_copy}"
        "【成文原则】\n"
        f"当前文种「{genre_profile.get('label') or template_id}」：{genre_profile.get('purpose', '')}。"
        "用户定制要求优先级最高；事实与数据只来自材料硬事实与用户明确写出的信息；"
        "体例 JSON 与范文样本提供篇章节奏与写法参照。"
        "任务是独立创作：提炼主题、设计结构、组织论述，用公文语体成文；"
        "不是压缩精简，也不是逐段粘贴原材料，更不是编造情节。\n"
        "主件写入成文主体，附件区长文仅在 [[attachment_head]] 点名。\n\n"
        "【创作方式】\n"
        "1. 先落实用户定制要求（侧重、语气、结构、增删意向）；\n"
        "2. 以结构映射计划为施工顺序，逐段按 writingMethod 创作（合并模式下可微调层次以利表达）；\n"
        "3. 以硬事实清单为成文主依据，原材料仅供核对与补全覆盖；"
        "须覆盖全部相关硬事实（数字、日期、专名、结论）且不得改写数值；\n"
        "4. 正文在材料事实上展开：设计过渡、强化逻辑、组织论述，句式须重写；"
        f"篇幅导向约 {creation_target} 字（约为材料字数 50%–60%，上限 {output_limit} 字），"
        "须明显短于原材料篇幅；\n"
        "5. 版头版脚（title/主送/落款/日期）由你据事实与文种体例判断取舍；"
        "主标题须独立润写概括事由，禁止照录材料原标题；"
        "主送、落款、日期等专名按事实保留；\n"
        "6. 每个 [[h1]] 后紧跟对应 [[body]]；块内可用「一是…二是…」并列。\n"
        f"{creation_guide}\n"
        "正向对照：\n"
        f"{creation_example}\n\n"
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
        "输出中的动作、要求、程序、责任与结果均能在主件材料或用户明确要求中找到依据；"
        "禁止编造模拟数据、虚构单位/日期/情节；"
        "以材料硬事实为成文依据，须改写表述与段落结构，不得照搬底稿原文；"
        "主标题据事由独立润写；材料若仅为章节小标题，则从会议/主题名称提炼 [[title]]，"
        "章节小标题保留为 [[h1]]。\n\n"
        "【输出格式】\n"
        "按结构映射计划与文种常用角色安排段落顺序；每段独立一行，必须以 [[role]] 开头。"
        "role 为 title、subtitle、recipient、body、h1、h2、h3、h4、"
        "attachment_head、attachment_other、sign_unit、sign_date、sign_contact、other。\n"
        "结构示例：\n"
        f"{example}"
        f"只输出上述结构化段落，不写说明或 Markdown。"
        "禁止输出英文分析、推理过程或「Need to analyze」类说明，必须直接以 [[title]]/[[body]] 等标记成文。"
        + (f"\n\n{rewrite_note}" if rewrite_note else "")
    )
    structure_json = json.dumps(structure_payload, ensure_ascii=False, indent=2)
    if omit_reference_body or rewrite_mode:
        reference_body = (
            "（体例坐标：参照段落职责、层次深度与写法，在材料事实上独立创作成文。）\n"
            f"{structure_json}"
        )
    else:
        reference_text = str(reference.get("text", "") or "")
        if len(reference_text) > AI_REFERENCE_SAMPLE_CHARS:
            reference_text = (
                reference_text[:AI_REFERENCE_SAMPLE_CHARS]
                + "\n…（范文样本已截取，体例以结构 JSON 为准）"
            )
        reference_body = (
            f"{structure_json}\n\n---\n体例样本原文（参照篇章骨架与写法，用材料事实独立创作）：\n"
            f"{reference_text}"
        )
    plan_label = (
        "结构映射计划（平台体例蓝图；合并模式下请在成文时内化整理与策划，按 writingMethod 创作）"
        if use_merge
        else "结构映射计划（成文蓝图，按 writingMethod 逐段创作）"
    )
    user = (
        f"{user_directive}\n\n"
        f"【目标文种锁定】必须严格按「{genre_profile.get('label') or template_id}」（{template_id or 'generic'}）成文，"
        f"篇章体例、标题称谓与范文参照均服从该文种，不得写成其他文种。\n"
        f"文种标识：{template_id or 'generic'}\n"
        f"自动匹配结果：\n{template_note}\n\n"
        + (f"{zones_notes}\n\n" if zones_notes else "")
        + f"{plan_label}：\n<plan>\n{plan_text}\n</plan>\n\n"
        f"模板结构化框架（JSON，文种体例坐标）：\n"
        f"<template_structure>\n{reference_body}\n</template_structure>\n\n"
        f"成文内容来源（①用户定制优先 ②硬事实清单 ③要点；默认不含原材料全文）：\n<source>\n"
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
    """照录风控：重新开一轮对话，只给事实要点，禁止参考上一稿与原材料原文。"""
    audit_warnings = list((reference_audit or {}).get("warnings") or [])
    source_coverage = (validation_metrics or {}).get("sourceCopyCoverage")
    body_coverage = (validation_metrics or {}).get("bodyCopyCoverage")
    copied_ratio = (validation_metrics or {}).get("copiedParagraphRatio")
    body_similarity = (validation_metrics or {}).get("bodySourceSimilarity")
    detail_parts = [str(error)]
    if body_similarity is not None:
        detail_parts.append(f"正文相似度={body_similarity}")
    if copied_ratio is not None:
        detail_parts.append(f"段落复制比={copied_ratio}")
    if body_coverage is not None:
        detail_parts.append(f"正文片段重合={body_coverage}")
    if source_coverage is not None:
        detail_parts.append(f"全文片段重合={source_coverage}")
    if audit_warnings:
        detail_parts.append("需规避表述：" + "；".join(audit_warnings[:3]))
    rewrite_reason = "；".join(detail_parts)
    messages = build_ai_messages(
        prompt, materials, template, reference, transformation_plan,
        omit_reference_body=True,
        rewrite_mode=True,
        rewrite_reason=rewrite_reason,
        merge_pipeline=AI_MERGE_PIPELINE,
    )
    # 不把上一稿塞回上下文，避免模型照抄自己；明确禁止复用原句式。
    messages.append({
        "role": "user",
        "content": (
            "【强制再创作·禁止照录】上一稿不合格：与原材料句式/段落近照录。"
            "请从零重写：仅依据事实清单中的硬事实，"
            "必须更换开头、过渡、段落顺序与句式（可用「一是…二是…」、"
            "「现将…函告如下」「现就…明确如下」等公文句式重组）；"
            "不得复述、改写幅度低于30%视为失败；"
            "主标题须据事实润写，正文必须重写。"
            "直接输出完整 [[role]] 结构化段落，不要解释。"
        ),
    })
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


def _join_message_content_parts(content: Any) -> str:
    if content is None:
        return ""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: list[str] = []
        for part in content:
            if isinstance(part, str):
                parts.append(part)
            elif isinstance(part, dict):
                parts.append(str(
                    part.get("text")
                    or part.get("content")
                    or part.get("reasoning")
                    or part.get("thinking")
                    or ""
                ))
            else:
                parts.append(str(part))
        return "".join(parts)
    return str(content)


def call_chat_model_result(
    request_url: str,
    api_key: str,
    model_name: str,
    messages: list[dict[str, str]],
    temperature: float = 0.2,
    max_tokens: int | None = None,
    *,
    thinking: str | None = None,
    response_format: dict[str, Any] | None = None,
    reasoning_effort: str | None = None,
    allow_reasoning_fallback: bool = True,
) -> dict[str, Any]:
    """调用 OpenAI 兼容 Chat Completions。

    thinking: None=网关默认；\"disabled\"/\"enabled\" 显式开关（deepseek-v4 等默认开启思考）。
    reasoning_effort: 可选 none/low/medium/high，部分网关用此关闭思考。
    """
    load_local_env()
    request_url = request_url.strip() or os.getenv("DEEPSEEK_BASE_URL", "")
    api_key = api_key.strip() or os.getenv("DEEPSEEK_API_KEY", "")
    model_name = model_name.strip() or os.getenv("DEEPSEEK_MODEL", "")
    if not model_name:
        raise ValueError("请填写模型名称")
    payload: dict[str, Any] = {
        "model": model_name,
        "messages": messages,
        "temperature": temperature,
        "stream": False,
    }
    if max_tokens:
        payload["max_tokens"] = max_tokens
    if thinking in {"disabled", "enabled"}:
        payload["thinking"] = {"type": thinking}
    if reasoning_effort:
        payload["reasoning_effort"] = reasoning_effort
    if response_format:
        payload["response_format"] = response_format
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    def _post(body_payload: dict[str, Any]) -> dict[str, Any]:
        raw = json.dumps(body_payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            resolve_chat_url(request_url), data=raw, headers=headers, method="POST",
        )
        with urllib.request.urlopen(request, timeout=MODEL_HTTP_TIMEOUT_SECONDS) as resp:
            return json.loads(resp.read().decode("utf-8"))

    try:
        body = _post(payload)
    except TimeoutError as exc:
        raise RuntimeError(
            f"模型接口超时（{MODEL_HTTP_TIMEOUT_SECONDS} 秒）。"
            "高并发时推理排队较长，可提高环境变量 AI_MODEL_HTTP_TIMEOUT_SECONDS 后重试"
        ) from exc
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        # 部分网关不认 response_format / thinking：按优先级去掉后重试，尽量保留 thinking=disabled
        if exc.code in {400, 422}:
            retry_payload = dict(payload)
            retried = False
            if "response_format" in retry_payload:
                retry_payload.pop("response_format", None)
                try:
                    body = _post(retry_payload)
                    retried = True
                except Exception:
                    retried = False
            if not retried and "reasoning_effort" in retry_payload:
                retry_payload.pop("reasoning_effort", None)
                try:
                    body = _post(retry_payload)
                    retried = True
                except Exception:
                    retried = False
            if not retried and "thinking" in retry_payload:
                retry_payload.pop("thinking", None)
                try:
                    body = _post(retry_payload)
                    retried = True
                except Exception:
                    retried = False
            if not retried:
                raise RuntimeError(f"模型接口返回 {exc.code}: {detail[:500]}") from exc
        else:
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
        content = _join_message_content_parts(message.get("content")).strip()
        reasoning = _join_message_content_parts(
            message.get("reasoning_content")
            or message.get("reasoning")
            or message.get("thinking")
            or ""
        ).strip()
        refusal = str(message.get("refusal") or "").strip()
        finish_reason = str(choice.get("finish_reason") or body.get("finish_reason") or "")
        # 思考模型常把可用结论放在 reasoning；成文等结构化输出默认不拿推理链顶替正文。
        if not content and reasoning and allow_reasoning_fallback:
            content = reasoning
        if not content and refusal:
            content = refusal
        return {
            "content": content,
            "reasoning": reasoning,
            "finishReason": finish_reason or "unknown",
            "model": body.get("model", model_name),
            "usage": body.get("usage") or {},
            "reasoningChars": len(reasoning),
            "rawChoice": {
                "hasContent": bool(_join_message_content_parts(message.get("content")).strip()),
                "hasReasoning": bool(reasoning),
                "hasRefusal": bool(refusal),
            },
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
        if RE_STANDALONE_PLACEHOLDER_LINE.match(stripped) or (
            RE_PLACEHOLDER_TEXT.search(stripped)
            and len(stripped) <= 24
            and not re.search(r"[\u4e00-\u9fff]{2,}", stripped.replace("请补充", "").replace("待补", "").replace("待定", "").replace("待填写", "").replace("请填写", ""))
        ):
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


def _looks_like_reasoning_meta_draft(text: str) -> bool:
    sample = str(text or "").strip()
    if not sample:
        return False
    if draft_has_role_markers(sample):
        return False
    if RE_REASONING_META_DRAFT.search(sample[:3000]):
        return True
    cjk = len(re.findall(r"[\u4e00-\u9fff]", sample[:3000]))
    ascii_letters = len(re.findall(r"[A-Za-z]", sample[:3000]))
    return ascii_letters >= 80 and cjk < 40


def _draft_attempt_budgets(speed_mode: str) -> list[int]:
    budgets = list(DRAFT_TOKEN_BUDGETS.get(speed_mode, DRAFT_TOKEN_BUDGETS["standard"]))
    primary = max(budgets) if budgets else 12288
    expanded = min(max(primary + 4096, 16384), 20480)
    if expanded not in budgets:
        budgets.append(expanded)
    return budgets


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
    budgets = _draft_attempt_budgets(speed_mode)
    attempts: list[dict[str, Any]] = []
    for max_tokens in budgets:
        result = call_chat_model_result(
            request_url, api_key, model_name, messages,
            temperature=temperature, max_tokens=max_tokens,
            thinking="disabled",
            reasoning_effort="none",
            allow_reasoning_fallback=False,
        )
        content = result.get("content") or ""
        usage = result.get("usage") or {}
        reasoning_tokens = int((usage.get("completion_tokens_details") or {}).get("reasoning_tokens") or 0)
        attempt = {
            "maxTokens": max_tokens,
            "finishReason": result.get("finishReason") or "unknown",
            "outputChars": len(content),
            "reasoningChars": result.get("reasoningChars", 0),
            "reasoningTokens": reasoning_tokens,
            "usage": usage,
        }
        attempts.append(attempt)
        if not content or RE_EMPTY_SUMMARY.search(content):
            if content:
                attempt["invalidOutput"] = "错误声明材料为空"
            elif reasoning_tokens >= max(int(max_tokens * 0.8), 1024):
                attempt["invalidOutput"] = "思考链占满输出额度，未返回可用正文"
            continue
        if _looks_like_reasoning_meta_draft(content):
            attempt["invalidOutput"] = "返回内容为英文分析/推理，非结构化公文"
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
            "thinkingDisabled": True,
        }
    details = "；".join(
        f"第 {index + 1} 次 max_tokens={item['maxTokens']}，finish_reason={item['finishReason']}，"
        f"正文={item['outputChars']} 字，推理={item['usage'].get('completion_tokens_details', {}).get('reasoning_tokens', 0)} tokens"
        f"{('，' + item['invalidOutput']) if item.get('invalidOutput') else ''}"
        for index, item in enumerate(attempts)
    )
    raise RuntimeError(
        f"模型成文未返回正文：{details}。"
        "请检查模型服务是否支持关闭思考（thinking=disabled），或改用可稳定输出正文的模型。"
    )


def call_transformation_planner(
    request_url: str, api_key: str, model_name: str,
    messages: list[dict[str, str]], speed_mode: str,
    materials: list[dict[str, Any]],
    reference: dict[str, Any] | None = None,
    template_id: str = "",
) -> tuple[dict[str, Any], dict[str, Any]]:
    budgets = DRAFT_TOKEN_BUDGETS.get(speed_mode, DRAFT_TOKEN_BUDGETS["standard"])
    max_tokens = min(max(budgets), 4096)
    resolved_template_id = template_id or str(
        (reference or {}).get("templateId")
        or ((reference or {}).get("structureJson") or {}).get("templateId")
        or ""
    )
    platform_plan = build_fallback_transformation_plan(resolved_template_id, materials, reference)
    if AI_MERGE_PIPELINE or AI_SKIP_PLAN_LLM:
        reason = (
            "结构策划已并入成文合并调用，使用文种/范文平台蓝图作为成文施工底图"
            if AI_MERGE_PIPELINE
            else "已跳过结构策划模型，使用范文/文种平台蓝图以加速一次成文"
        )
        return platform_plan, {
            "maxTokens": 0,
            "outputChars": 0,
            "usage": {},
            "finishReason": "merged" if AI_MERGE_PIPELINE else "skipped",
            "reasoningChars": 0,
            "fallback": True,
            "fallbackReason": reason,
            "mergedIntoDraft": bool(AI_MERGE_PIPELINE),
        }
    meta: dict[str, Any] = {"maxTokens": max_tokens}
    try:
        result = call_chat_model_result(
            request_url, api_key, model_name, messages,
            temperature=0.0, max_tokens=max_tokens,
        )
    except Exception as exc:
        meta.update({
            "fallback": True,
            "fallbackReason": f"结构策划模型调用失败：{exc}",
            "outputChars": 0,
            "finishReason": "error",
            "reasoningChars": 0,
            "usage": {},
        })
        return platform_plan, meta

    content = result.get("content") or ""
    meta.update({
        "outputChars": len(content),
        "usage": result.get("usage") or {},
        "finishReason": result.get("finishReason") or "unknown",
        "reasoningChars": result.get("reasoningChars") or 0,
    })
    if not content:
        meta["fallback"] = True
        meta["fallbackReason"] = (
            f"结构策划模型未返回正文（finish_reason={meta['finishReason']}，"
            f"reasoning={meta['reasoningChars']} 字），已用平台体例蓝图继续成文"
        )
        return platform_plan, meta
    try:
        plan = parse_transformation_plan(content, materials)
        if "title" not in {item.get("role") for item in plan.get("targetSections") or []}:
            plan["targetSections"] = [
                *platform_plan["targetSections"][:1],
                *plan["targetSections"],
            ]
        return plan, meta
    except RuntimeError as exc:
        meta["fallback"] = True
        meta["fallbackReason"] = str(exc)
        return platform_plan, meta


def validate_reference_fact_isolation(
    draft: str, reference: dict[str, Any], prompt: str, materials: list[dict[str, Any]],
    strict: bool | None = None,
) -> dict[str, Any]:
    """审计范文事实隔离；strict=True 时检测到泄漏则阻断成文。"""
    source_text = "\n".join([
        prompt,
        *(item.get("rawText") or item.get("text", "") for item in materials),
    ])
    source_facts = set(RE_REFERENCE_FACT_TOKEN.findall(source_text))
    reference_facts = set(RE_REFERENCE_FACT_TOKEN.findall(reference.get("text", "")))
    output_facts = set(RE_REFERENCE_FACT_TOKEN.findall(draft))
    leaked = sorted(
        (reference_facts - source_facts) & output_facts,
        key=lambda item: (-len(item), item),
    )
    # 过滤过短数字（如 15、19、21），避免范文页码/序号误报为事实串入
    leaked = [
        token for token in leaked
        if not (token.isdigit() and len(token) < 4)
    ]
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


def _blocks_text_by_roles(blocks: list[dict[str, str]], roles: frozenset[str]) -> str:
    return "\n".join(
        str(block.get("text") or "").strip()
        for block in blocks
        if block.get("role") in roles and str(block.get("text") or "").strip()
    ).strip()


def _enumerate_cn(index: int) -> str:
    chars = "一二三四五六七八九十"
    if 1 <= index <= len(chars):
        return chars[index - 1]
    return str(index)


def restructure_copied_body_blocks(
    blocks: list[dict[str, str]], materials: list[dict[str, Any]],
) -> tuple[list[dict[str, str]], list[str]]:
    """照录兜底：对仍与材料段落高度重合的正文做规则化重组。"""
    source = "\n".join(
        item.get("rawText") or item.get("text", "") for item in materials
    )
    source_paragraphs: list[str] = []
    for line in source.replace("\r\n", "\n").split("\n"):
        text = line.strip()
        if len(text) >= 16:
            source_paragraphs.append(text)
    for part in re.split(r"\n\s*\n", source):
        text = part.strip()
        if len(text) >= 16 and text not in source_paragraphs:
            source_paragraphs.append(text)
    if not source_paragraphs:
        return blocks, []

    result = [dict(block) for block in blocks]
    warnings: list[str] = []
    for block in result:
        role = block.get("role")
        if role not in BODY_DRAFT_ROLES:
            continue
        text = str(block.get("text") or "").strip()
        if len(text) < 24:
            continue
        best_ratio = max(
            (SequenceMatcher(None, text, paragraph).ratio() for paragraph in source_paragraphs),
            default=0.0,
        )
        if best_ratio < 0.82:
            continue
        sentences = [
            part.strip()
            for part in re.split(r"(?<=[。；;])", text)
            if part.strip() and len(part.strip()) >= 6
        ]
        if len(sentences) < 2:
            sentences = [text]
        rephrased: list[str] = []
        for index, sentence in enumerate(sentences, start=1):
            core = sentence.rstrip("。；;")
            if index == 1:
                rephrased.append(f"一是{core}。")
            elif index == 2:
                rephrased.append(f"二是{core}。")
            else:
                rephrased.append(f"{_enumerate_cn(index)}是{core}。")
        block["text"] = "".join(rephrased)
        warnings.append(f"已对近照录正文段（{role}）做规则化重组")
    return result, warnings


def _placeholder_hits_not_in_source(draft: str, materials: list[dict[str, Any]], prompt: str = "") -> list[str]:
    """找出成稿中的待填占位；原材料/用户要求里已有的表述不算占位。"""
    source = "\n".join([
        prompt or "",
        *(str(item.get("rawText") or item.get("text") or "") for item in materials),
    ])
    hits: list[str] = []
    for match in RE_PLACEHOLDER_TEXT.finditer(draft or ""):
        token = match.group(0)
        if token and token in source:
            continue
        hits.append(token)
    return hits


def validate_transformed_draft(
    blocks: list[dict[str, str]], materials: list[dict[str, Any]],
    transformation_plan: dict[str, Any],
) -> dict[str, Any]:
    """评估成文质量。

    创作成文与原材料事实接近属正常；
    仅在大段近照录（段落级 SequenceMatcher）时触发重构。
    占位符不再硬拦整链：材料原有脱敏名（如张某某）放行，真正的待填标记只记警告。
    """
    draft = "\n".join(block.get("text", "") for block in blocks).strip()
    body_draft = _blocks_text_by_roles(blocks, BODY_DRAFT_ROLES)
    warnings: list[str] = []
    placeholder_hits = _placeholder_hits_not_in_source(draft, materials)
    if placeholder_hits:
        warnings.append(
            "成稿含待填标记（已放行平台整理，请人工核验）："
            + "、".join(list(dict.fromkeys(placeholder_hits))[:6])
        )

    source = "\n".join(
        item.get("rawText") or item.get("text", "") for item in materials
    )
    brief_source = "\n".join(
        item.get("rawText") or item.get("text", "") for item in materials
    )
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
    body_copied_chars = 0
    source_paragraphs: list[str] = []
    for part in re.split(r"\n\s*\n", brief_source or source):
        text = part.strip()
        if len(text) >= 20:
            source_paragraphs.append(text)
    for line in (brief_source or source).replace("\r\n", "\n").split("\n"):
        text = line.strip()
        if len(text) >= 20 and text not in source_paragraphs:
            source_paragraphs.append(text)
    for block in blocks:
        text = block.get("text", "").strip()
        role = block.get("role")
        if len(text) < 20:
            continue
        best_ratio = max(
            (SequenceMatcher(None, text, paragraph).ratio() for paragraph in source_paragraphs),
            default=0.0,
        )
        if best_ratio >= 0.82:
            copied_chars += len(text)
            if role in BODY_DRAFT_ROLES:
                body_copied_chars += len(text)
    copied_ratio = copied_chars / max(len(draft), 1)
    body_copied_ratio = body_copied_chars / max(len(body_draft), 1) if body_draft else 0.0
    draft_source_similarity = round(
        SequenceMatcher(
            None,
            re.sub(r"\s+", "", draft)[:12000],
            re.sub(r"\s+", "", source)[:12000],
        ).ratio(),
        4,
    )
    body_source_similarity = round(
        SequenceMatcher(
            None,
            re.sub(r"\s+", "", body_draft)[:12000],
            re.sub(r"\s+", "", source)[:12000],
        ).ratio(),
        4,
    ) if body_draft else 0.0
    body_shingles = _normalized_shingles(body_draft)
    body_copy_coverage = (
        len(source_shingles & body_shingles) / len(body_shingles)
        if body_shingles else 0.0
    )
    expected_roles = [
        str(section.get("role"))
        for section in transformation_plan.get("targetSections", [])
        if section.get("role") != "other"
    ]
    actual_roles = [block.get("role") for block in blocks if block.get("role") != "other"]
    structure_coverage = _lcs_length(expected_roles, actual_roles) / max(len(expected_roles), 1)
    # 照录判定以正文段落/正文整体相似度为主；版头版脚原字保留不计入致命重合。
    similarity_risk = bool(
        AI_REWRITE_ON_MATERIAL_SIMILARITY
        and (
            body_copied_ratio >= SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD
            or copied_ratio >= SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD
            or body_source_similarity >= SIMILARITY_BODY_SOURCE_THRESHOLD
            or (body_draft and body_copy_coverage >= SIMILARITY_REWRITE_COPY_THRESHOLD)
        )
    )
    if body_copied_ratio >= SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD or copied_ratio >= SIMILARITY_REWRITE_PARAGRAPH_THRESHOLD:
        if similarity_risk:
            warnings.append(
                "生成结果存在大段近照录材料段落，将自动再创作（禁止照录）"
            )
        else:
            warnings.append(
                f"成文与材料段落重合较高（正文 {body_copied_ratio:.2f}），已放行（未开启照录再创作）"
            )
    elif body_source_similarity >= SIMILARITY_BODY_SOURCE_THRESHOLD:
        if similarity_risk:
            warnings.append(
                f"正文与原材料整体相似度过高（{body_source_similarity:.2f}），将自动再创作（禁止照录）"
            )
        else:
            warnings.append(
                f"正文与原材料整体相似度过高（{body_source_similarity:.2f}），已放行"
            )
    elif draft_source_similarity >= SIMILARITY_DRAFT_SOURCE_THRESHOLD:
        if similarity_risk:
            warnings.append(
                f"成稿与原材料整体相似度过高（{draft_source_similarity:.2f}），将自动再创作（禁止照录）"
            )
        else:
            warnings.append(
                f"成稿与原材料整体相似度过高（{draft_source_similarity:.2f}），已放行"
            )
    elif source_copy_coverage >= SIMILARITY_REWRITE_COPY_THRESHOLD:
        if similarity_risk:
            warnings.append(
                f"成文与材料事实重合度较高（{source_copy_coverage:.2f}），将自动再创作（禁止照录）"
            )
        else:
            warnings.append(
                f"成文与材料事实重合度较高（{source_copy_coverage:.2f}），已放行"
            )
    if structure_coverage < 0.6:
        warnings.append("成文角色序列与参考范文差异较大，建议人工核验段落组织")
    return {
        "sourceCopyCoverage": round(source_copy_coverage, 4),
        "briefCopyCoverage": round(brief_copy_coverage, 4),
        "copiedParagraphRatio": round(copied_ratio, 4),
        "bodyCopiedParagraphRatio": round(body_copied_ratio, 4),
        "draftSourceSimilarity": draft_source_similarity,
        "bodySourceSimilarity": body_source_similarity,
        "bodyCopyCoverage": round(body_copy_coverage, 4),
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
            if value:
                blocks.append({"role": role, "text": value})
                if RE_PLACEHOLDER_TEXT.search(value) and len(value) <= 24:
                    warnings.append("段落含待填标记，已保留供平台整理")
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


GENRE_TITLE_MARKERS: dict[str, tuple[str, ...]] = {
    "order": ("命令", "令"),
    "announcement": ("公告",),
    "notice": ("通知",),
    "bulletin": ("通报",),
    "letter": ("函",),
    "public_notice": ("通告",),
    "decision": ("决定",),
    "resolution": ("决议",),
    "report": ("报告", "汇报"),
    "request": ("请示",),
    "opinion": ("意见",),
    "minutes": ("纪要",),
}


def _material_title_conflicts_with_genre(title: str, template_id: str) -> bool:
    text = str(title or "").strip()
    if not text or not template_id:
        return False
    expected = GENRE_TITLE_MARKERS.get(template_id, ())
    if expected and any(marker in text for marker in expected):
        return False
    for other_id, markers in GENRE_TITLE_MARKERS.items():
        if other_id == template_id:
            continue
        if any(marker in text for marker in markers):
            return True
    return False


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
        item.get("summary") or item.get("rawText") or item.get("text", "")
        for item in materials
    ).strip()
    facts = [
        f"{fact.get('id')}:{fact.get('content')}"
        for item in materials
        for fact in (item.get("structuredFacts") or [])
        if fact.get("content")
    ]

    # 先规则补标题；默认不再为补齐字段额外打模型，避免二次排队。
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

    if remaining and AI_SKIP_FILL_LLM:
        warnings.append(
            f"已跳过缺失字段模型补齐（{len(remaining)} 项），保留现有段落加速一次成文"
        )
        remaining = []

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
                    f"材料事实摘要：\n{material_brief}\n"
                    f"事实条目：\n{chr(10).join(facts) or '无'}\n\n"
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
    template_id: str = "",
) -> tuple[list[dict[str, str]], list[str]]:
    """补回模型遗漏的公文要素；不覆盖模型已写内容，标题与附件说明由模型润写。"""
    result = [dict(block) for block in blocks]
    warnings: list[str] = []
    model_roles = {block.get("role") for block in result}
    # 标题、副标题、附件说明交由成文模型判断与润写，不强制贴回材料原文
    skip_roles = {"title", "subtitle", "attachment_head"}
    required = [
        element
        for material in materials
        for element in material.get("sourceElements", [])
    ]
    for element in required:
        role, text = element["role"], element["text"]
        if role in skip_roles:
            continue
        if role == "recipient" and template_id == "order":
            if role not in model_roles:
                warnings.append("命令（令）文种不设主送机关，已跳过材料主送行")
            continue
        if role in model_roles:
            continue
        if role == "sign_unit":
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
            warnings.append("模型未写落款单位，已据材料补全")
            model_roles.add("sign_unit")
            continue
        if role in {"subtitle", "recipient", "security"}:
            header_roles = {"title", "subtitle", "recipient", "security"}
            insert_at = sum(1 for block in result if block.get("role") in header_roles)
            result.insert(insert_at, {"role": role, "text": text})
        else:
            result.append({"role": role, "text": text})
        warnings.append(f"模型未写{ROLE_NAMES.get(role, role)}，已据材料补全")
        model_roles.add(role)
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
    template_key: str | None,
    request_url: str,
    api_key: str,
    model_name: str,
    temperature: float,
    speed_mode: str,
    strict_reference_isolation: bool = False,
    allow_degradation: bool = False,
    format_output: bool = True,
) -> dict[str, Any]:
    resolved_template_id, template = resolve_generation_template(
        prompt,
        [material["text"]],
        template_id=template_id,
        template_key=template_key,
    )
    if not template:
        raise RuntimeError(f"未找到目标公文模板：{resolved_template_id}")
    warnings: list[str] = list(material.get("warnings") or [])
    template_label = template["label"] if template else resolved_template_id
    prepared_material, material_warnings = prepare_material(
        path=material.get("path"),
        prompt=prompt,
        template_label=template_label,
        request_url=request_url,
        api_key=api_key,
        model_name=model_name,
        template_id=resolved_template_id,
        inventory=material.get("inventory"),
        name=material.get("name"),
        allow_degradation=allow_degradation,
    )
    current_materials = [prepared_material]
    warnings.extend(material_warnings)
    if prepared_material.get("processingMode") == "rule_based_brief" and not AI_SKIP_PREP_LLM:
        warnings.append(
            "材料整理阶段已回退为平台本地抽取；成文依据规则化事实清单"
        )
    reference = select_reference_for_material(template, prompt, current_materials)
    if reference.get("warning"):
        warnings.append(f"参考范文解析提示：{reference['warning']}")
    plan_started_at = perf_counter()
    if AI_MERGE_PIPELINE or AI_SKIP_PLAN_LLM:
        transformation_plan, planner_meta = call_transformation_planner(
            request_url, api_key, model_name, [], speed_mode, current_materials,
            reference=reference, template_id=resolved_template_id,
        )
    else:
        plan_messages = build_transformation_plan_messages(
            prompt, current_materials, template, reference,
        )
        transformation_plan, planner_meta = call_transformation_planner(
            request_url, api_key, model_name, plan_messages, speed_mode, current_materials,
            reference=reference, template_id=resolved_template_id,
        )
    if planner_meta.get("mergedIntoDraft") or AI_MERGE_PIPELINE:
        warnings.append(
            "结构策划已并入成文一次调用：整理模型已输出事实清单，"
            "成文时按体例蓝图独立创作"
        )
    elif planner_meta.get("fallback") and not AI_SKIP_PLAN_LLM:
        warnings.append(
            "结构策划采用平台体例蓝图："
            f"{planner_meta.get('fallbackReason', '未知原因')}"
        )
    elif AI_SKIP_PLAN_LLM:
        warnings.append("结构策划走平台蓝图快路径，跳过策划模型以一次成文")
    plan_seconds = round(perf_counter() - plan_started_at, 2)
    messages = build_ai_messages(
        prompt, current_materials, template, reference, transformation_plan,
        merge_pipeline=AI_MERGE_PIPELINE,
    )
    drafting_started_at = perf_counter()
    rewrite_retries = 0
    similarity_rewrites = 0
    copy_rewrites = 0
    targeted_fill_meta: dict[str, Any] = {"filledRoles": [], "mode": "noop"}
    validation_metrics: dict[str, Any] = {}
    last_error: RuntimeError | None = None
    last_reference_audit: dict[str, Any] | None = None
    draft = ""
    generation_meta: dict[str, Any] = {}
    blocks: list[dict[str, str]] = []
    base_messages = list(messages)
    guard_limit = max(MAX_GENERATION_GUARD_ATTEMPTS, MAX_COPY_REWRITE_ATTEMPTS + 1)
    for guard_attempt in range(guard_limit):
        if last_error is not None:
            if is_similarity_failure(last_error):
                messages = build_similarity_rewrite_messages(
                    prompt, current_materials, template, reference, transformation_plan,
                    last_error,
                    reference_audit=last_reference_audit,
                    validation_metrics=validation_metrics,
                )
                similarity_rewrites += 1
                copy_rewrites += 1
                draft_temperature = min(temperature + 0.12 * copy_rewrites, 0.55)
            else:
                messages = [
                    *messages,
                    {"role": "assistant", "content": draft},
                    {"role": "user", "content": build_parse_retry_message(last_error)},
                ]
                draft_temperature = temperature
                rewrite_retries += 1
        else:
            messages = base_messages
            draft_temperature = temperature
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
            skip_header_preserve = copy_rewrites > 0
            if skip_header_preserve:
                preservation_warnings = []
                warnings.append("照录再创作轮次：仅校验正文，不再强制贴回材料版头/附件说明")
            else:
                blocks, preservation_warnings = preserve_source_elements(
                    blocks, current_materials, resolved_template_id,
                )
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
                    title = derive_title_from_materials(current_materials, prompt, template_label)
                    blocks = _insert_filled_block(blocks, "title", title)
                    warnings.append(f"定向补齐后仍缺标题，已用材料主题补写：{title}")
            blocks = normalize_document_header_order(blocks)

            validation_metrics = validate_transformed_draft(
                blocks, current_materials, transformation_plan,
            )
            if validation_metrics.get("similarityRisk") and copy_rewrites >= MAX_COPY_REWRITE_ATTEMPTS:
                blocks, restructure_warnings = restructure_copied_body_blocks(blocks, current_materials)
                warnings.extend(restructure_warnings)
                validation_metrics = validate_transformed_draft(
                    blocks, current_materials, transformation_plan,
                )
            validation_metrics["referenceIsolation"] = reference_audit
            validation_metrics["targetedFill"] = targeted_fill_meta
            warnings.extend(validation_metrics.get("warnings") or [])
            similarity_risk = bool(
                reference_audit.get("isolationRisk")
                or (
                    AI_REWRITE_ON_MATERIAL_SIMILARITY
                    and validation_metrics.get("similarityRisk")
                )
            )
            if similarity_risk:
                reason_parts = []
                if reference_audit.get("isolationRisk"):
                    reason_parts.append("范文表述串入")
                if validation_metrics.get("similarityRisk"):
                    reason_parts.append("材料正文近照录")
                if copy_rewrites < MAX_COPY_REWRITE_ATTEMPTS:
                    last_error = RuntimeError(
                        "成文与材料原文重合度偏高，禁止照录："
                        + "、".join(reason_parts or ["正文近照录"])
                        + "，须改写句式后重新输出"
                    )
                    continue
                body_sim = validation_metrics.get("bodySourceSimilarity")
                body_copy = validation_metrics.get("bodyCopiedParagraphRatio")
                raise RuntimeError(
                    f"成文与材料近照录，已自动再创作 {copy_rewrites} 次仍不达标"
                    f"（正文相似度 {body_sim}，段落照录比 {body_copy}）。"
                    "请补充写作要求（如侧重、语气、结构）后重试，或缩短材料后生成。"
                )
            if similarity_rewrites:
                warnings.append(
                    f"已通过禁止照录再创作 {similarity_rewrites} 次，完成正文差异化"
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
    generation_meta["copyRewrites"] = copy_rewrites
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
        "templateKey": build_template_key(template),
        "templateName": template["name"] if template else "",
        "reference": {
            "filename": reference["filename"],
            "similarity": reference["similarity"],
            "textChars": reference["textChars"],
        },
        "transformationPlan": transformation_plan,
        "preview": preview,
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
            "prepSeconds": current_materials[0].get("prepSeconds"),
            "planSeconds": plan_seconds,
            "draftingSeconds": round(perf_counter() - drafting_started_at, 2),
            "rendering": "材料整理 + 结构策划 + 创作成文",
        },
    }


def generate_documents(
    prompt: str,
    upload_paths: list[Path],
    request_url: str,
    api_key: str,
    model_name: str,
    template_id: str | None = None,
    template_key: str | None = None,
    temperature: float = 0.2,
    speed_mode: str = "standard",
    strict_reference_isolation: bool = False,
    allow_degradation: bool = False,
    format_output: bool = True,
    text_materials: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    started_at = perf_counter()
    if speed_mode not in {"fast", "standard", "deep"}:
        speed_mode = "standard"
    resolved_template_id = str(template_id or "").strip()
    if template_key:
        resolved_template_id, _ = resolve_generation_template(
            prompt, [], template_id=template_id, template_key=template_key,
        )
    elif resolved_template_id:
        resolved_template_id = normalize_template_id(resolved_template_id)
    if not resolved_template_id or resolved_template_id == "generic":
        raise RuntimeError("请选择目标公文模板（通知、报告、函等）后再生成")
    text_materials = [
        item for item in (text_materials or [])
        if str(item.get("text") or "").strip()
    ]
    if not upload_paths and not text_materials:
        raise RuntimeError("请上传材料文件，或在对话框填写写作要求/粘贴文本材料后再生成")
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

    for index, item in enumerate(text_materials, start=1):
        extraction_started_at = perf_counter()
        text = str(item.get("text") or "").strip()
        name = str(item.get("name") or f"对话框材料{index}.txt").strip() or f"对话框材料{index}.txt"
        inventory = build_plain_text_inventory(text)
        material_quality = assess_material_content(inventory)
        if material_quality["status"] != "ready":
            failures.append({
                "filename": name,
                "reason": f"对话框材料正文不足：{material_quality['reason']}",
            })
            continue
        targets.append({
            "path": None,
            "name": name,
            "text": text,
            "warnings": ["材料来自对话框文本，已跳过文件提取"],
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
                template_id=resolved_template_id,
                template_key=template_key,
                request_url=request_url,
                api_key=api_key,
                model_name=model_name,
                temperature=temperature,
                speed_mode=speed_mode,
                strict_reference_isolation=strict_reference_isolation,
                allow_degradation=allow_degradation,
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
        "batchHint": "每个上传材料或对话框文本材料单独生成一份公文，不会自动汇总成一篇。",
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
