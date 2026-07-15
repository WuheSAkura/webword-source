"""Word 公文格式转换引擎 —— 拆分 + 多信号识别（带置信度）+ 按校正结果套用。

流程：
  build_structure(path)            识别每段类型 + 置信度，供人工校正
  convert_with_roles(in,out,roles) 按用户校正后的类型套格式，返回转换日志
所有格式参数来自 config.get_config()，不在本文件硬编码。
"""

import re
import copy as _copy
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from config import get_config

ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "justify": 3}
CHAR_WIDTH_CM = 0.565  # 三号字单字宽近似，仅用于无 *Chars 属性时的兜底

# 低置信度阈值：低于此值前端高亮提醒人工复核
LOW_CONFIDENCE = 0.7

# ── 识别正则 ──
RE_H1 = re.compile(r"^[一二三四五六七八九十百]+\s*[、.．]")
RE_H2 = re.compile(r"^[（(]\s*[一二三四五六七八九十百]+\s*[)）]")
RE_H4 = re.compile(r"^[（(]\s*\d+\s*[)）]")
RE_H3 = re.compile(r"^\d+\s*[.．、](?=\s*\S)")
RE_ATTACH = re.compile(r"^附件\s*\d*\s*[:：]?")
RE_RECIPIENT = re.compile(r"[：:]\s*$")  # 主送机关（发文对象）：以冒号结尾的称呼行
RE_DATE = re.compile(r"\d{4}\s*年\s*\d{0,2}\s*月\s*\d{0,2}\s*日")
RE_DATE_FULL = re.compile(r"^\s*\d{4}\s*年\s*\d{0,2}\s*月\s*\d{0,2}\s*日\s*$")
RE_SECURITY = re.compile(r"^(绝密|机密|秘密|内部)(★|▲)?")
RE_YEAR_HEAD = re.compile(r"^\d{4}\s*年")  # 排除“2026年…”被误判为三级标题
RE_FULL_BRACKET = re.compile(r"^[（(].+[)）]$")

# 文种关键词（用于标题启发式）
TITLE_END = ("通知", "通报", "报告", "请示", "方案", "意见", "决定", "函", "规定",
             "办法", "条例", "细则", "纪要", "批复", "公告", "通告", "命令", "决议",
             "汇报", "计划", "总结", "说明", "情况", "意见稿")
RE_TITLE_KW = re.compile(
    r"(关于|关于印发|转发).{0,80}?"
    r"(通知|通报|报告|请示|方案|意见|决定|函|规定|办法|条例|细则|纪要|批复|公告|通告|命令|决议|汇报|计划|总结|说明|情况)"
)
SUBTITLE_HINT = ("单位", "职务", "部", "局", "厅", "委", "办", "处", "科", "院", "司",
                 "组", "室", "中心", "主任", "局长", "部长", "处长", "科长", "书记")
QUOTE_CHARS = "“”‘’\"'《》〈〉"

# ── 段落拆分（沿用：把混排编号拆成逻辑段落）──
RE_SPLIT_MARKER = re.compile(
    r'(?P<h1>[一二三四五六七八九十]{1,3}[、，,])'
    r'|(?P<h2>[（(][一二三四五六七八九十]{1,3}[）)])'
    r'|(?P<h4>[（(]\d{1,2}[）)])'
    r'|(?P<h3>(?<!\d)\d{1,2}[.．、](?=\s*\S))'
    r'|(?P<attach>附件\s*[：:\d])'
)


def is_valid_marker(text: str, match: re.Match) -> bool:
    """过滤示例/引文里的编号，避免把说明性文字拆碎。"""
    start, end = match.start(), match.end()
    prev = text[start - 1] if start > 0 else ""
    next_char = text[end] if end < len(text) else ""
    if prev in QUOTE_CHARS or next_char in QUOTE_CHARS:
        return False
    if prev == "第" and match.lastgroup == "h1":
        return False
    if match.lastgroup == "h3":
        if prev.isdigit() or (prev.isascii() and prev.isalpha()) or prev in "+-±.":
            return False
        if next_char.isdigit():
            return False
        if prev in "：:" or "附件" in text[max(0, start - 4):start]:
            return False
    if match.lastgroup == "attach" and prev in (QUOTE_CHARS + "如见"):
        return False
    return True


def split_paragraph_text(text: str) -> list[str]:
    """将一个段落文本按编号分隔符拆分为多个逻辑段落"""
    text = re.sub(r"\s*\n+\s*", "\n", text.strip())
    if not text:
        return [text]
    # 整段被括号包裹（副标题、落款联系人等）不拆分，
    # 避免内部数字/顿号（如“张三，”“电话xxx”）被误当作编号切断
    if RE_FULL_BRACKET.match(text):
        return [text]
    markers = [m for m in RE_SPLIT_MARKER.finditer(text) if is_valid_marker(text, m)]
    if not markers:
        return refine_logical_parts([text])
    parts = []
    first_start = markers[0].start()
    if first_start > 0:
        before = text[:first_start].strip()
        if before:
            parts.append(before)
    for idx, marker in enumerate(markers):
        start = marker.start()
        next_start = markers[idx + 1].start() if idx + 1 < len(markers) else len(text)
        segment = text[start:next_start].strip()
        if segment:
            parts.append(segment)
    return refine_logical_parts(parts)


def split_prelude_title(text: str) -> list[str]:
    """从无换行的文头里拆出主标题，避免文号/主送机关吞掉标题。"""
    if RE_H1.match(text) or RE_H2.match(text) or RE_H3.match(text):
        return [text]
    match = RE_TITLE_KW.search(text)
    if not match:
        return [text]
    title = match.group(0).strip()
    if not (4 <= len(title) <= 90):
        return [text]
    before = text[:match.start()].strip()
    after = text[match.end():].strip()
    result = []
    if before:
        result.extend([p for p in re.split(r"\s{2,}|\n+", before) if p.strip()])
    result.append(title)
    if after:
        result.append(after)
    return result


def split_heading_body(text: str) -> list[str]:
    """把编号标题后紧跟的正文拆开，避免整段正文套标题字体。"""
    t = text.strip()
    if not t:
        return []
    if RE_H1.match(t):
        match = re.search(r"\s+", t)
        if match:
            heading, body = t[:match.start()].strip(), t[match.end():].strip()
            if 4 <= len(heading) <= 30 and body:
                return [heading, body]
        return [t]
    h2_match = RE_H2.match(t)
    h3_match = RE_H3.match(t)
    if h2_match or h3_match:
        search_start = h3_match.end() if h3_match else 0
        match = re.search(r"[。.!！?？]", t[search_start:])
        if match and match.end() < len(t):
            end = search_start + match.end()
            heading, body = t[:end].strip(), t[end:].strip()
            if 4 <= len(heading) <= 45 and body:
                return [heading, body]
        return [t]
    return [t]


def refine_logical_parts(parts: list[str]) -> list[str]:
    refined = []
    for part in parts:
        for prelude_part in split_prelude_title(part.strip()):
            refined.extend(split_heading_body(prelude_part))
    return [p for p in refined if p.strip()]


# ── XML 段落操作 ──
def insert_paragraph_after(para_element):
    new_p = OxmlElement("w:p")
    para_element.addnext(new_p)
    return new_p


def insert_paragraph_before(para_element):
    new_p = OxmlElement("w:p")
    para_element.addprevious(new_p)
    return new_p


def split_document_paragraphs(doc: Document) -> list:
    """遍历所有段落，对含多个编号块的段落进行拆分。返回 (element, text) 列表。"""
    result = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            result.append((para._element, text))
            continue
        parts = split_paragraph_text(text)
        if len(parts) <= 1:
            result.append((para._element, text))
            continue
        para.text = ""
        if parts[0].strip():
            para.add_run(parts[0])
        result.append((para._element, parts[0]))
        elem = para._element
        for part in parts[1:]:
            if not part.strip():
                continue
            elem = insert_paragraph_after(elem)
            r_elem = OxmlElement("w:r")
            t_elem = OxmlElement("w:t")
            t_elem.text = part.strip()
            t_elem.set(qn("xml:space"), "preserve")
            r_elem.append(t_elem)
            elem.append(r_elem)
            result.append((elem, part.strip()))
    return result


# ── 多信号识别 ──
def _para_format_hint(para) -> dict:
    """提取段落原格式线索：是否居中、最大字号、是否加粗。"""
    centered = para.alignment == 1
    max_size = 0.0
    bold = False
    for run in para.runs:
        try:
            if run.font.size:
                max_size = max(max_size, run.font.size.pt)
        except Exception:
            pass
        if run.font.bold:
            bold = True
    return {"centered": centered, "max_size": max_size, "bold": bold}


def classify_one(text: str, index: int, total: int, fmt: dict | None = None) -> tuple[str, float]:
    """对单个逻辑段落做基础分类，返回 (role, confidence)。多信号：正则+位置+原格式+长度。"""
    t = text.strip()
    fmt = fmt or {}
    if not t:
        return "body", 0.0

    near_head = index <= 2
    near_tail = (total - index) <= 5

    # 1. 强编号特征（顺序：H2/H4 先于 H3，避免“（1）”被当 H3）
    if RE_H1.match(t):
        return "h1", 0.95
    if RE_H2.match(t):
        return "h2", 0.95
    if RE_H4.match(t):
        return "h4", 0.95
    if RE_ATTACH.match(t):
        return "attachment_head", 0.95
    if RE_H3.match(t):
        if RE_YEAR_HEAD.match(t):  # 排除“2026年…”
            return "body", 0.6
        # 三级标题 vs 编号列表正文：过长且以句号结尾 → 低置信，交人工
        if len(t) > 40 and t[-1] in "。.":
            return "h3", 0.5
        return "h3", 0.9

    # 2. 涉密标识（可选扩展类型）
    if index <= 1 and RE_SECURITY.match(t) and len(t) <= 30:
        return "security", 0.85

    # 3. 成文日期（落款区）
    if RE_DATE_FULL.match(t) and near_tail:
        return "sign_date", 0.95

    # 4. 标题（文头区）
    if near_head and 4 <= len(t) <= 90:
        if RE_TITLE_KW.match(t) or (len(t) <= 35 and any(t.endswith(k) for k in TITLE_END)):
            return "title", 0.9
        # 原格式线索：居中 + 大字号 + 加粗 → 强指向标题
        if fmt.get("centered") and fmt.get("max_size", 0) >= 18:
            return "title", 0.75 if fmt.get("bold") else 0.7

    # 5. 副标题：文头整段括号文本
    if near_head and RE_FULL_BRACKET.match(t) and len(t) <= 60:
        return "subtitle", 0.75

    # 5b. 主送机关（发文对象）：文头区、以冒号结尾的称呼行，顶格不缩进
    #     排除“……如下：”等正文引出句，避免误判
    if (index <= 6 and RE_RECIPIENT.search(t) and 2 <= len(t) <= 60
            and not RE_DATE.search(t) and "。" not in t
            and not t.endswith(("如下：", "如下:"))):
        return "recipient", 0.85

    # 5c. 落款·联系人：文末括号行，含“联系人/电话”，首行缩进2字（不与日期对齐）
    if near_tail and RE_FULL_BRACKET.match(t) and ("联系" in t or "电话" in t):
        return "sign_contact", 0.85

    # 6. 落款单位：文末短行、非长句
    if near_tail and len(t) <= 30 and t[-1] not in "。.！？!?；;，,、" and not RE_DATE.search(t):
        return "sign_unit", 0.6

    return "body", 0.8


def refine_structure(items: list[dict]) -> None:
    """二次校正：利用相邻关系修正副标题/落款配对。就地修改 items。"""
    n = len(items)
    # 发文对象之前基本都是标题：把误判为正文/编号的段落回归标题
    # （标题可跨多行，但行与行之间不应被识别成正文，避免套格式时夹空行）
    recip_idx = next((i for i, it in enumerate(items) if it["role"] == "recipient"), None)
    if recip_idx is not None:
        for i in range(recip_idx):
            if items[i]["role"] in ("body", "h1", "h2", "h3", "h4", "other"):
                items[i]["role"] = "title"
                items[i]["confidence"] = max(items[i]["confidence"], 0.9)
    # 副标题必须紧跟标题之后
    for i in range(1, n):
        if items[i]["role"] == "subtitle" and items[i - 1]["role"] != "title":
            # 上一段不是标题：若本段不像括号副标题则回退正文
            items[i]["role"] = "body"
            items[i]["confidence"] = 0.6
    # 多个附件并列：紧跟“附件头”之后、形如“2.×××”的编号行归为“其他附件”
    # （可链式：其他附件后面的编号行仍是其他附件）
    for i in range(1, n):
        if items[i - 1]["role"] not in ("attachment_head", "attachment_other"):
            continue
        t = items[i]["fullText"].strip()
        if (items[i]["role"] in ("h3", "body") and RE_H3.match(t)
                and t[-1:] not in "。.！？!?；;"):
            items[i]["role"] = "attachment_other"
            items[i]["confidence"] = max(items[i]["confidence"], 0.7)
    # 落款单位应紧邻成文日期（其上一段）
    date_idx = [i for i, it in enumerate(items) if it["role"] == "sign_date"]
    if date_idx:
        di = date_idx[0]
        if di - 1 >= 0 and items[di - 1]["role"] == "body":
            prev = items[di - 1]
            if len(prev["fullText"]) <= 30 and prev["fullText"][-1:] not in "。.！？!?":
                prev["role"] = "sign_unit"
                prev["confidence"] = max(prev["confidence"], 0.65)


def build_structure(file_path: str) -> list[dict]:
    """识别文档结构，返回带置信度的清单（供人工校正与套格式共用）。"""
    config = get_config()
    doc = Document(str(file_path))
    para_items = split_document_paragraphs(doc)

    # 建立 element → Paragraph 映射以读原格式
    para_by_elem = {p._element: p for p in doc.paragraphs}

    logical = [(elem, txt) for elem, txt in para_items if txt.strip()]
    total = len(logical)
    items = []
    for i, (elem, txt) in enumerate(logical):
        para = para_by_elem.get(elem)
        fmt = _para_format_hint(para) if para is not None else {}
        role, conf = classify_one(txt, i, total, fmt)
        items.append({
            "index": i,
            "role": role,
            "confidence": round(conf, 2),
            "lowConfidence": conf < LOW_CONFIDENCE,
            "text": txt[:40] + ("…" if len(txt) > 40 else ""),
            "fullText": txt,
        })
    refine_structure(items)
    for it in items:
        it["lowConfidence"] = it["confidence"] < LOW_CONFIDENCE
    return items


# ── 格式应用 ──
def nuke_run_format(rPr):
    if rPr is None:
        return
    tags = ("w:rFonts", "w:sz", "w:szCs", "w:b", "w:bCs", "w:i", "w:iCs",
            "w:u", "w:color", "w:spacing", "w:kern", "w:position", "w:vertAlign",
            "w:strike", "w:dstrike", "w:shadow", "w:outline", "w:emboss",
            "w:imprint", "w:highlight", "w:lang", "w:fitText", "w:effect",
            "w:w", "w:snapToGrid", "w:shd", "w:smallCaps", "w:caps", "w:vanish",
            "w:webHidden", "w:specVanish", "w:em", "w:cs")
    for tag in tags:
        e = rPr.find(qn(tag))
        if e is not None:
            rPr.remove(e)


def set_font(run, east: str, west: str, size: float, bold: bool = False):
    """一个 run 同时设 ascii/hAnsi/cs=西文、eastAsia=中文，Word 按字符自动选字体。"""
    rPr = run._element.get_or_add_rPr()
    nuke_run_format(rPr)
    run.font.name = west
    run.font.size = Pt(size)
    run.font.bold = bold
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), east)
    rFonts.set(qn("w:ascii"), west)
    rFonts.set(qn("w:hAnsi"), west)
    rFonts.set(qn("w:cs"), west)


def _set_ind_attr(paragraph, attr: str, value: str):
    # 用 get_or_add_ind 保证 w:ind 在 pPr 中的 schema 顺序正确
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(qn(attr), value)


def set_first_line_chars(paragraph, chars: int):
    if chars and chars > 0:
        _set_ind_attr(paragraph, "w:firstLineChars", str(int(chars * 100)))


def set_right_chars(paragraph, chars: float):
    if chars and chars > 0:
        _set_ind_attr(paragraph, "w:rightChars", str(int(round(chars * 100))))


def set_left_chars(paragraph, chars: float):
    if chars and chars > 0:
        _set_ind_attr(paragraph, "w:leftChars", str(int(round(chars * 100))))


def style_for(role: str, config: dict) -> dict:
    """从配置取出某 role 的内部样式字典。"""
    styles = config["styles"]
    s = styles.get(role, styles["body"])
    line = s.get("line_pt")
    if line is None:
        line = config["spacing"]["title_line_pt"] if role == "title" else config["spacing"]["body_line_pt"]
    return {
        "role": role,
        "font_east": s["cn"],
        "font_west": config["fonts"]["en"],
        "size": s["size_pt"],
        "bold": s.get("bold", False),
        "align": s.get("align", "justify"),
        "line_spacing": line,
        "first_line_chars": s.get("first_line_chars", 0),
    }


def apply_paragraph_layout(paragraph, style: dict, right_chars_override: float | None = None,
                           left_chars_override: float | None = None):
    align = style.get("align", "justify")
    fmt = paragraph.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.line_spacing = Pt(style.get("line_spacing", 28))
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.left_indent = None
    fmt.right_indent = None
    fmt.first_line_indent = None
    # 清掉旧的 ind，避免残留
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is not None:
        old_ind = pPr.find(qn("w:ind"))
        if old_ind is not None:
            pPr.remove(old_ind)

    if align == "center_to_date":
        paragraph.alignment = ALIGN_MAP["right"]
        if right_chars_override is not None:
            set_right_chars(paragraph, right_chars_override)
    elif align == "right_indent_4chars":
        paragraph.alignment = ALIGN_MAP["right"]
        set_right_chars(paragraph, 4)
    elif align == "right_indent_2chars":
        # 落款单位/日期：右对齐，离右边距 2 个汉字（右端对齐，不用空格）
        paragraph.alignment = ALIGN_MAP["right"]
        set_right_chars(paragraph, 2)
    else:
        paragraph.alignment = ALIGN_MAP.get(align, 3)
        set_first_line_chars(paragraph, style.get("first_line_chars", 0))
        # 多个附件时，第二个起按"附件："宽度做悬挂缩进，使序号上下对齐
        if left_chars_override is not None:
            set_left_chars(paragraph, left_chars_override)


def apply_style(paragraph, style: dict, right_chars_override: float | None = None,
                left_chars_override: float | None = None):
    apply_paragraph_layout(paragraph, style, right_chars_override, left_chars_override)
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is not None:
        p_rPr = pPr.find(qn("w:rPr"))
        if p_rPr is not None:
            nuke_run_format(p_rPr)
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        set_font(run, style["font_east"], style["font_west"], style["size"], style["bold"])


def clean_text(text: str) -> str:
    text = text.replace("　", "")
    text = re.sub(r"[ \t ]+", " ", text)
    text = re.sub(r"^([一二三四五六七八九十百]+[、.．])\s+", r"\1", text)
    text = re.sub(r"^([（(][一二三四五六七八九十百]+[）)])\s+", r"\1", text)
    text = re.sub(r"^([（(]\d+[）)])\s+", r"\1", text)
    text = re.sub(r"^(\d+[.．、])\s+", r"\1", text)
    return text.strip()


# ── 页面 / 空行 / 页码 ──
def set_page(doc, page: dict):
    margin = page.get("margin_cm", {})
    for sec in doc.sections:
        sec.top_margin = Cm(margin.get("top", 3.7))
        sec.bottom_margin = Cm(margin.get("bottom", 3.5))
        sec.left_margin = Cm(margin.get("left", 2.8))
        sec.right_margin = Cm(margin.get("right", 2.6))
        sec.header_distance = Cm(page.get("header_cm", 0))
        sec.footer_distance = Cm(page.get("footer_cm", 2.5))
        sec.page_width = Cm(page.get("width_cm", 21.0))
        sec.page_height = Cm(page.get("height_cm", 29.7))


def paragraph_element_text(elem) -> str:
    return "".join(t.text or "" for t in elem.findall(".//" + qn("w:t")))


def ensure_blank_paragraphs_before(para_element, count: int):
    existing = 0
    prev = para_element.getprevious()
    while prev is not None and prev.tag == qn("w:p"):
        if paragraph_element_text(prev).strip():
            break
        existing += 1
        prev = prev.getprevious()
    for _ in range(max(count - existing, 0)):
        insert_paragraph_before(para_element)


def remove_blank_paragraphs_before(para_element):
    """删除紧邻 para_element 之前的所有空段落（用于标题各行之间不留空行）。"""
    prev = para_element.getprevious()
    while prev is not None and prev.tag == qn("w:p") and not paragraph_element_text(prev).strip():
        to_remove = prev
        prev = prev.getprevious()
        to_remove.getparent().remove(to_remove)


def _enable_even_odd_headers(doc: Document):
    settings = doc.settings.element
    if settings.find(qn("w:evenAndOddHeaders")) is None:
        el = OxmlElement("w:evenAndOddHeaders")
        settings.append(el)


def add_page_number_field(paragraph, align_key: str, page_number: dict):
    paragraph.text = ""
    paragraph.alignment = ALIGN_MAP.get(align_key, 1)
    fmt_str = page_number.get("format", "—{PAGE}—")
    prefix, _, suffix = fmt_str.partition("{PAGE}")
    cn = page_number.get("cn", "宋体")
    size = page_number.get("size_pt", 14)
    west = "Times New Roman"

    if prefix:
        r = paragraph.add_run(prefix)
        set_font(r, cn, west, size)

    r = paragraph.add_run()
    fb = OxmlElement("w:fldChar")
    fb.set(qn("w:fldCharType"), "begin")
    r._r.append(fb)
    r = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    r._r.append(instr)
    set_font(r, cn, west, size)
    r = paragraph.add_run()
    fe = OxmlElement("w:fldChar")
    fe.set(qn("w:fldCharType"), "end")
    r._r.append(fe)

    if suffix:
        r = paragraph.add_run(suffix)
        set_font(r, cn, west, size)


def set_page_numbers(doc: Document, page_number: dict):
    odd_right = page_number.get("odd_right_even_left", True)
    if odd_right:
        _enable_even_odd_headers(doc)
    for sec in doc.sections:
        footer = sec.footer
        para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        add_page_number_field(para, "right" if odd_right else "center", page_number)
        if odd_right:
            even = sec.even_page_footer
            epara = even.paragraphs[0] if even.paragraphs else even.add_paragraph()
            add_page_number_field(epara, "left", page_number)


# ── 主流程：按校正后的角色套格式 ──
def _display_width(text: str) -> float:
    """粗略显示宽度：中文记 1，数字/字母/标点记 0.5。用于落款居中计算。"""
    w = 0.0
    for ch in text:
        w += 0.5 if ch.isascii() else 1.0
    return w


def convert_with_roles(input_path: str, output_path: str,
                       roles: dict[int, str] | None = None,
                       clean_spaces: bool = True) -> dict:
    """按 roles（{逻辑段索引: 角色}）套格式；缺失的段落回退自动识别。返回转换日志。"""
    config = get_config()
    roles = {int(k): v for k, v in (roles or {}).items()}

    doc = Document(str(input_path))
    para_items = split_document_paragraphs(doc)
    para_by_elem = {p._element: p for p in doc.paragraphs}

    logical = [(elem, txt) for elem, txt in para_items if txt.strip()]
    total = len(logical)

    # 先算好每段最终角色（含自动识别兜底），便于跨段落计算（落款居中）
    final_roles = []
    fmt_cache = []
    for i, (elem, txt) in enumerate(logical):
        para = para_by_elem.get(elem)
        fmt = _para_format_hint(para) if para is not None else {}
        fmt_cache.append(fmt)
        auto_role, _ = classify_one(txt, i, total, fmt)
        role = roles.get(i, auto_role)
        if role not in config["styles"]:
            role = "body"
        final_roles.append(role)

    blank = config["blank_lines"]
    seen_title = False
    seen_attachment = False
    seen_signature = False
    log = []

    for i, (elem, txt) in enumerate(logical):
        para = para_by_elem.get(elem)
        if para is None:
            continue
        role = final_roles[i]
        style = style_for(role, config)

        # 空行规范化
        if role == "title":
            if not seen_title:
                # 只在标题首行上方留空行
                ensure_blank_paragraphs_before(elem, blank.get("before_title", 2))
                seen_title = True
            else:
                # 标题其余各行：删掉行间空行，只换行不空行
                remove_blank_paragraphs_before(elem)
        elif role == "attachment_head" and not seen_attachment:
            # 附件区与正文之间空行，只在第一段“附件头”前加
            ensure_blank_paragraphs_before(elem, blank.get("before_attachment", 1))
            seen_attachment = True
        elif role == "sign_unit" and not seen_signature:
            ensure_blank_paragraphs_before(elem, blank.get("before_signature", 3))
            seen_signature = True

        cleaned = clean_text(txt) if clean_spaces else txt
        if cleaned != para.text and para.runs:
            runs_len = sum(len(r.text) for r in para.runs)
            for r in para.runs:
                r.text = ""
            if para.runs and runs_len > 0:
                para.runs[0].text = cleaned

        apply_style(para, style)
        log.append({"index": i, "role": role,
                    "corrected": i in roles and roles[i] != classify_one(txt, i, total, fmt_cache[i])[0],
                    "text": txt[:30]})

    set_page(doc, config["page"])
    set_page_numbers(doc, config["page_number"])

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return {"output": output_path, "log": log}


def convert_docx(input_path: str, output_path: str, clean_spaces: bool = True) -> str:
    """无人工校正的一键转换（自动识别 + 套格式），用于兼容/兜底。"""
    result = convert_with_roles(input_path, output_path, roles=None, clean_spaces=clean_spaces)
    return result["output"]


# ── 局部字体微调（套格式后的补充编辑）──
def clear_runs(paragraph):
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)


def apply_inline_font(run, font: dict, fallback: dict):
    east = font.get("font_east") or fallback["font_east"]
    west = font.get("font_west") or fallback["font_west"]
    size = font.get("size") or fallback["size"]
    bold = font.get("bold")
    if bold is None:
        bold = fallback["bold"]
    set_font(run, east, west, size, bold)
    if font.get("italic") is not None:
        run.font.italic = bool(font.get("italic"))
    if font.get("underline") is not None:
        run.font.underline = bool(font.get("underline"))


def build_style(role: str | None, overrides: dict | None = None) -> dict:
    config = get_config()
    base_role = role if role in config["styles"] else "body"
    style = style_for(base_role, config)
    overrides = overrides or {}
    if overrides.get("alignment"):
        style["align"] = overrides["alignment"]
    if overrides.get("line_spacing") is not None:
        style["line_spacing"] = overrides["line_spacing"]
    if overrides.get("first_indent") is not None:
        style["first_line_chars"] = overrides["first_indent"]
    return style


def rewrite_paragraph_with_selection(paragraph, start, end, base_style, font):
    text = paragraph.text
    start = max(0, min(start, len(text)))
    end = max(start, min(end, len(text)))
    before, selected, after = text[:start], text[start:end], text[end:]
    clear_runs(paragraph)
    for segment, is_sel in ((before, False), (selected, True), (after, False)):
        if not segment:
            continue
        run = paragraph.add_run(segment)
        if is_sel and font:
            apply_inline_font(run, font, base_style)
        else:
            set_font(run, base_style["font_east"], base_style["font_west"],
                     base_style["size"], base_style["bold"])


def apply_local_edit(docx_path: str, selection: dict, font: dict | None = None,
                     paragraph: dict | None = None) -> str:
    doc = Document(str(docx_path))
    paragraphs = doc.paragraphs
    start_para = int(selection.get("startParagraph", 0))
    end_para = int(selection.get("endParagraph", start_para))
    start_offset = int(selection.get("startOffset", 0))
    end_offset = int(selection.get("endOffset", 0))
    if start_para > end_para:
        start_para, end_para = end_para, start_para
        start_offset, end_offset = end_offset, start_offset
    start_para = max(0, min(start_para, len(paragraphs) - 1))
    end_para = max(0, min(end_para, len(paragraphs) - 1))
    non_empty = [p.text for p in paragraphs if p.text.strip()]
    total = len(non_empty)

    paragraph = paragraph or {}
    role_override = paragraph.get("role") or None

    for para_index in range(start_para, end_para + 1):
        para = paragraphs[para_index]
        text = para.text
        if not text:
            continue
        role = role_override or classify_one(text, para_index, total)[0]
        base_style = build_style(role, paragraph)
        apply_paragraph_layout(para, base_style)
        seg_start = start_offset if para_index == start_para else 0
        seg_end = end_offset if para_index == end_para else len(text)
        if seg_end < seg_start:
            seg_start, seg_end = seg_end, seg_start
        rewrite_paragraph_with_selection(para, seg_start, seg_end, base_style, font)

    doc.save(str(docx_path))
    return docx_path
