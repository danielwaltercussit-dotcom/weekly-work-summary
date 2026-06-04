# -*- coding: utf-8 -*-
"""周报 docx 排版样式助手。

集中管理字体、配色、徽章、分隔线、页边距、页脚等底层 OOXML 操作,
让 build_docx.py 专注于内容结构。所有颜色用 6 位十六进制(不带 #)。
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Cm

# ---- 字体 ----
FONT_TITLE_EA = "微软雅黑"      # 标题 / 小节标题:略重,撑得起层次
FONT_BODY_EA = "等线"          # 正文:更轻,长文阅读更舒服
FONT_LATIN = "Segoe UI"        # 拉丁字母与数字:日期、298 这类更清爽

# ---- 配色 ----
PRIMARY = "1F4E79"     # 主题深蓝:标题、小节标题
ACCENT = "2E75B6"      # 强调中蓝:横幅、小节左侧色条
BODY = "262626"        # 正文近黑,比纯黑柔和
MUTED = "7F7F7F"       # 次要灰:日期副标题、明细说明
HAIRLINE = "D9D9D9"    # 细分隔线浅灰

# 状态徽章:浅底 + 深字,清楚但不刺眼
BADGE = {
    "done":    {"label": "已完成", "fill": "E6F4EA", "text": "1E7E34"},
    "doing":   {"label": "进行中", "fill": "E7F0FA", "text": "1565C0"},
    "blocked": {"label": "受阻",   "fill": "FDECEA", "text": "C62828"},
}


def _rpr(run):
    return run._element.get_or_add_rPr()


def style_run(run, *, latin=FONT_LATIN, ea=FONT_BODY_EA, size=None,
              bold=False, italic=False, color=None):
    """给一个 run 设字体(中西文分别指定)、字号、加粗、颜色。"""
    rpr = _rpr(run)
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:cs"), latin)
    rFonts.set(qn("w:eastAsia"), ea)
    run.font.name = latin
    if size is not None:
        run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def run_shading(run, fill_hex):
    """给 run 加底色(徽章效果)。"""
    rpr = _rpr(run)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    rpr.append(shd)


def add_badge(paragraph, status):
    """在段落开头插入一个带底色的状态徽章 run。"""
    b = BADGE.get(status)
    if not b:
        return
    run = paragraph.add_run(" " + b["label"] + " ")  # 窄空格当内边距
    style_run(run, ea=FONT_TITLE_EA, size=9, bold=True, color=b["text"])
    run_shading(run, b["fill"])
    sep = paragraph.add_run("  ")
    style_run(sep, size=10)


def _border_el(tag, sz, color, space=4):
    el = OxmlElement(tag)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))      # 1/8 pt:8=1pt,24=3pt
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    return el


def paragraph_borders(paragraph, *, bottom=None, left=None):
    """给段落加边框。bottom/left 各为 (sz, color) 或 None。"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    if left is not None:
        pBdr.append(_border_el("w:left", left[0], left[1], space=6))
    if bottom is not None:
        pBdr.append(_border_el("w:bottom", bottom[0], bottom[1], space=4))


def set_normal_style(doc):
    """设全局 Normal 样式:中西文字体 + 行距 + 段后距,确保未单独设样式的文字也好看。"""
    normal = doc.styles["Normal"]
    normal.font.name = FONT_LATIN
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(BODY)
    rpr = normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_LATIN)
    rFonts.set(qn("w:hAnsi"), FONT_LATIN)
    rFonts.set(qn("w:eastAsia"), FONT_BODY_EA)
    pf = normal.paragraph_format
    pf.line_spacing = 1.4
    pf.space_after = Pt(6)


def set_margins(section, top=2.4, bottom=2.4, left=2.6, right=2.6):
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)


def add_footer_pagenum(section, left_text):
    """页脚:左侧自定义文字,右侧页码。"""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.text = ""
    r = p.add_run(left_text)
    style_run(r, size=8.5, color=MUTED)
    # 用制表位把页码推到右侧
    tab = p.add_run("\t\t")
    style_run(tab, size=8.5)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    fr = OxmlElement("w:r")
    ft = OxmlElement("w:t")
    ft.text = "1"
    fr.append(ft)
    fld.append(fr)
    p._p.append(fld)
    pr = p.add_run("")
    style_run(pr, size=8.5, color=MUTED)
