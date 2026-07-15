"""公文格式配置加载层 —— 从 gongwen_config.yaml 读取唯一参数源。

对外暴露 get_config()，返回带兜底默认值的配置字典；YAML 缺失或损坏时不崩。
"""

import copy
from pathlib import Path

import yaml

CONFIG_PATH = Path(__file__).resolve().parent / "gongwen_config.yaml"

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
    "spacing": {"body_line_pt": 28, "title_line_pt": 32},
    "blank_lines": {"before_title": 2, "before_attachment": 1, "before_signature": 3},
    "styles": {
        "title":      {"cn": "方正小标宋简体", "size_pt": 22, "bold": False, "align": "center",  "first_line_chars": 0, "line_pt": 32},
        "subtitle":   {"cn": "楷体_GB2312",   "size_pt": 16, "bold": False, "align": "center",  "first_line_chars": 0},
        "recipient":  {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 0},
        "body":       {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h1":         {"cn": "黑体",          "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h2":         {"cn": "楷体_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h3":         {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
        "h4":         {"cn": "仿宋_GB2312",   "size_pt": 16, "bold": False, "align": "justify", "first_line_chars": 2},
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
    loaded = {}
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
