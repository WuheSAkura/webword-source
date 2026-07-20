"""公文格式配置加载层 —— 从 gongwen_config.yaml 读取唯一参数源。

对外暴露 get_config()，返回带兜底默认值的配置字典；YAML 缺失或损坏时不崩。
"""

import copy
from pathlib import Path

import yaml

CONFIG_PATH = Path(__file__).resolve().parent / "gongwen_config.yaml"
TEMPLATE_CONFIG_PATH = Path(__file__).resolve().parent / "document_templates.yaml"

# ── 兜底默认值（YAML 缺失时使用，与 yaml 保持一致）──
DEFAULT_CONFIG = {
    "fonts": {
        "cn_fallback": {
            "title": ["方正小标宋简体", "宋体", "SimSun"],
            "fangsong": ["仿宋_GB2312", "仿宋", "FangSong"],
            "kaiti": ["楷体_GB2312", "楷体", "KaiTi"],
            "heiti": ["黑体", "SimHei"],
            "songti": ["宋体", "SimSun"],
        },
        "en": "Times New Roman",
    },
    "page": {
        "width_cm": 21.0,
        "height_cm": 29.7,
        "margin_cm": {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
        "header_cm": 0,
        "footer_cm": 2.5,
    },
    "spacing": {
        "body_line_rule": "exact", "body_line_pt": 28, "body_line_multiple": 1.0,
        "title_line_rule": "exact", "title_line_pt": 32, "title_line_multiple": 1.0,
    },
    "blank_lines": {"before_title": 2, "before_attachment": 1, "before_signature": 3},
    "inline_hierarchy": {
        "prefer_source_style_boundary": True,
        "terminators": "。！？；：.!?;:",
        "body_markers": ["一是", "二是", "三是", "四是", "五是", "六是", "七是", "八是", "九是", "十是"],
        "body_marker_bold": True,
    },
    "styles": {
        "title":      {"cn": "方正小标宋简体", "size_pt": 22, "bold": False, "align": "center",  "first_line_chars": 0, "line_pt": 32},
        "subtitle":   {"cn": "楷体_GB2312",   "size_pt": 16, "bold": False, "align": "center",  "first_line_chars": 0},
        "recipient":  {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 0},
        "body":       {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h1":         {"cn": "黑体",          "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h2":         {"cn": "楷体_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h3":         {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": True,  "align": "justify", "first_line_chars": 2},
        "h4":         {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": True,  "align": "justify", "first_line_chars": 2},
        "attachment_head":  {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "attachment_other": {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 5},
        "sign_unit":  {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "right_indent_2chars", "first_line_chars": 0},
        "sign_date":  {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "right_indent_2chars", "first_line_chars": 0},
        "sign_contact": {"cn": "仿宋_GB2312", "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "security":   {"cn": "黑体",          "size_pt": 16, "bold": False, "align": "left",    "first_line_chars": 0},
    },
    "page_number": {
        "cn": "宋体",
        "size_pt": 14,
        "format": "— {PAGE} —",
        "odd_right_even_left": True,
    },
}

# 中文显示标签（前端共用，避免两处重复维护）
ROLE_LABELS = {
    "title": "标题",
    "subtitle": "副标题",
    "recipient": "发文对象",
    "body": "正文",
    "h1": "一级标题",
    "h2": "二级标题",
    "h3": "三级标题",
    "h4": "四级标题",
    "attachment_head": "附件头",
    "attachment_other": "其他附件",
    "sign_unit": "落款·单位",
    "sign_date": "落款·日期",
    "sign_contact": "落款·联系人",
    "security": "涉密标识",
    "other": "其他",
    "blank": "空行",
}

_cache: dict | None = None
_template_cache: dict | None = None

DEFAULT_TEMPLATE_CONFIG = {
    "default": "generic",
    "templates": {
        "generic": {
            "order": 0,
            "label": "通用转换",
            "description": "使用现有通用规则识别并套用公文格式",
            "classification": {
                "title_keywords": [],
                "recipient": "optional",
                "subtitle": "optional",
                "signature": "optional",
            },
            "format": {},
        }
    },
}


def _deep_merge(base: dict, override: dict) -> dict:
    """用 override 覆盖 base，递归合并字典。"""
    result = copy.deepcopy(base)
    for key, val in (override or {}).items():
        if isinstance(val, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], val)
        else:
            result[key] = val
    return result


def load_config(force: bool = False) -> dict:
    """加载配置，YAML 覆盖默认值。结果缓存，force=True 可强制重载。"""
    global _cache
    if _cache is not None and not force:
        return _cache
    loaded: dict = {}
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = yaml.safe_load(f) or {}
    except (yaml.YAMLError, OSError):
        loaded = {}
    _cache = _deep_merge(DEFAULT_CONFIG, loaded)
    return _cache


def get_config(force: bool = False) -> dict:
    return load_config(force)


def load_template_config(force: bool = False) -> dict:
    """加载文种模板目录；配置缺失时保留通用转换。"""
    global _template_cache
    if _template_cache is not None and not force:
        return _template_cache
    loaded: dict = {}
    try:
        if TEMPLATE_CONFIG_PATH.exists():
            with open(TEMPLATE_CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = yaml.safe_load(f) or {}
    except (yaml.YAMLError, OSError):
        loaded = {}
    _template_cache = _deep_merge(DEFAULT_TEMPLATE_CONFIG, loaded)
    return _template_cache


def get_default_template_id() -> str:
    template_config = load_template_config()
    template_id = str(template_config.get("default") or "generic")
    return template_id if template_id in template_config["templates"] else "generic"


def has_document_template(template_id: str | None) -> bool:
    if not template_id:
        return False
    return template_id in load_template_config()["templates"]


def normalize_template_id(template_id: str | None) -> str:
    return str(template_id) if has_document_template(template_id) else get_default_template_id()


def get_template_definition(template_id: str | None) -> dict:
    resolved_id = normalize_template_id(template_id)
    return load_template_config()["templates"][resolved_id]


def get_document_config(template_id: str | None = None) -> dict:
    """返回基础公文配置与指定文种覆盖项合并后的最终配置。"""
    definition = get_template_definition(template_id)
    return _deep_merge(get_config(), definition.get("format") or {})


def get_template_classification(template_id: str | None = None) -> dict:
    return copy.deepcopy(get_template_definition(template_id).get("classification") or {})


def get_resolved_styles(template_id: str | None = None) -> dict:
    """返回补齐有效行距规则后的样式，供前端微调面板使用。"""
    config = get_document_config(template_id)
    spacing = config["spacing"]
    styles = copy.deepcopy(config["styles"])
    for role, style in styles.items():
        prefix = "title" if role == "title" else "body"
        style.setdefault("line_rule", spacing.get(f"{prefix}_line_rule", "exact"))
        style.setdefault("line_pt", spacing.get(f"{prefix}_line_pt", 32 if prefix == "title" else 28))
        style.setdefault("line_multiple", spacing.get(f"{prefix}_line_multiple", 1.0))
    return styles


def get_template_catalog() -> list[dict]:
    """返回排序后的前端模板目录，不暴露内部覆盖实现。"""
    templates = load_template_config()["templates"]
    catalog = []
    for template_id, definition in templates.items():
        catalog.append({
            "id": template_id,
            "label": definition.get("label", template_id),
            "description": definition.get("description", ""),
            "order": int(definition.get("order", 999)),
            "styles": get_resolved_styles(template_id),
        })
    return sorted(catalog, key=lambda item: (item["order"], item["label"]))
