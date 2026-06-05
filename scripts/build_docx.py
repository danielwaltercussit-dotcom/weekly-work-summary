#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""根据周报内容 JSON 生成排版精致的 Word(.docx)周报。

用法:
    python build_docx.py --json content.json --out 周报_2026-06-03.docx

JSON 结构见 ../references/docx-content-schema.md。
排版样式(字体/配色/徽章/分隔线)集中在 docx_style.py。
依赖 python-docx(pip install python-docx)。
"""
import argparse
import json
import os
import sys

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_style as S

# Windows 控制台默认 GBK,打印中文路径会乱码;强制 stdout 用 UTF-8。
try:
    sys.stdout.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass


def _spacer(doc, pts=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run("")
    S.style_run(r, size=pts)
    return p


def add_banner(doc):
    """顶部一条主题色细横幅(用底色段落模拟),拉开标题与页边。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    S.paragraph_borders(p, bottom=(24, S.ACCENT))  # 3pt 中蓝下边框
    r = p.add_run("")
    S.style_run(r, size=2)


def add_title_block(doc, content):
    name = content.get("name", "").strip()
    start = content.get("period_start", "")
    end = content.get("period_end", "")

    add_banner(doc)

    # 主标题
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    r = p.add_run((name + " · " if name else "") + "工作周报")
    S.style_run(r, ea=S.FONT_TITLE_EA, size=22, bold=True, color=S.PRIMARY)

    # 副标题:日期范围
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(2)
    sr = sub.add_run("汇报周期　{} 至 {}".format(start, end))
    S.style_run(sr, size=10.5, color=S.MUTED)

    # 速览行:N 项成果 · M 项进行中 · K 项待支持
    n_done = len(content.get("achievements", []))
    n_doing = len(content.get("in_progress", []))
    n_block = len(content.get("blockers", []))
    bits = ["本周 {} 项成果".format(n_done)]
    if n_doing:
        bits.append("{} 项进行中".format(n_doing))
    if n_block:
        bits.append("{} 项待支持".format(n_block))
    ov = doc.add_paragraph()
    ov.paragraph_format.space_after = Pt(6)
    ovr = ov.add_run("　·　".join(bits))
    S.style_run(ovr, ea=S.FONT_TITLE_EA, size=10, bold=True, color=S.ACCENT)


def render_summary(doc, summary):
    summary = (summary or "").strip()
    if not summary:
        return
    add_section_heading(doc, "摘要")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(8)
    p.paragraph_format.right_indent = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    S.paragraph_borders(p, left=(18, S.HAIRLINE))
    r = p.add_run(summary)
    S.style_run(r, size=10.5, color="404040")


def add_section_heading(doc, text):
    """小节标题:左侧色条 + 底部细分隔线,撑出层次。"""
    _spacer(doc, 6)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    S.paragraph_borders(p, left=(28, S.ACCENT), bottom=(6, S.HAIRLINE))
    r = p.add_run(text)
    S.style_run(r, ea=S.FONT_TITLE_EA, size=14, bold=True, color=S.PRIMARY)
    return p


def add_subgroup_heading(doc, text):
    """项目分组小标题(二级)。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("▍" + text)
    S.style_run(r, ea=S.FONT_TITLE_EA, size=11.5, bold=True, color=S.ACCENT)
    return p


def add_bullet_title(doc, title, status=None):
    """成果/事项的标题行:实心圆点 + 可选徽章 + 加粗标题。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Pt(2)
    dot = p.add_run("●  ")
    S.style_run(dot, size=9, color=S.ACCENT)
    if status:
        S.add_badge(p, status)
    tr = p.add_run(title)
    S.style_run(tr, ea=S.FONT_TITLE_EA, size=11, bold=True, color=S.BODY)
    return p


def add_detail(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(20)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    S.style_run(r, size=10.5, color="404040")
    return p


def add_support_line(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(20)
    p.paragraph_format.space_after = Pt(4)
    tag = p.add_run("需要支持　")
    S.style_run(tag, ea=S.FONT_TITLE_EA, size=10, bold=True, color=S.BADGE["blocked"]["text"])
    r = p.add_run(text)
    S.style_run(r, size=10.5, color="404040", italic=True)
    return p


def render_achievements(doc, achievements):
    add_section_heading(doc, "本周核心成果")
    if not achievements:
        add_detail(doc, "本周暂无明确成果记录。")
        return
    projects = [a.get("project") for a in achievements if a.get("project")]
    distinct = list(dict.fromkeys(projects))
    grouped = len(distinct) > 1
    if grouped:
        for proj in distinct:
            add_subgroup_heading(doc, proj)
            for a in achievements:
                if a.get("project") == proj:
                    _one(doc, a)
        rest = [a for a in achievements if not a.get("project")]
        if rest:
            add_subgroup_heading(doc, "其他")
            for a in rest:
                _one(doc, a)
    else:
        for a in achievements:
            _one(doc, a)


def _one(doc, a):
    add_bullet_title(doc, a.get("title", ""), a.get("status"))
    if a.get("detail"):
        add_detail(doc, a["detail"])


def render_in_progress(doc, items):
    if not items:
        return
    add_section_heading(doc, "进行中事项")
    for it in items:
        title = it.get("title", "")
        if it.get("eta"):
            title += "（预计 {} 完成）".format(it["eta"])
        add_bullet_title(doc, title, "doing")
        if it.get("detail"):
            add_detail(doc, it["detail"])


def render_blockers(doc, blockers):
    add_section_heading(doc, "问题与需要的支持")
    if not blockers:
        add_detail(doc, "本周无需上级协调的阻塞事项。")
        return
    for b in blockers:
        add_bullet_title(doc, b.get("title", ""), "blocked")
        if b.get("detail"):
            add_detail(doc, b["detail"])
        if b.get("need"):
            add_support_line(doc, b["need"])


def render_next_week(doc, items):
    add_section_heading(doc, "下周计划")
    if not items:
        add_detail(doc, "待定。")
        return
    for i, it in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Pt(2)
        num = p.add_run("{}　".format(i))
        S.style_run(num, ea=S.FONT_TITLE_EA, size=11, bold=True, color=S.ACCENT)
        r = p.add_run(it)
        S.style_run(r, size=10.5, color=S.BODY)


def build(content, out_path):
    doc = Document()
    S.set_normal_style(doc)
    S.set_margins(doc.sections[0])

    name = content.get("name", "").strip()
    footer_text = "{}工作周报".format(name + " · " if name else "")
    S.add_footer_pagenum(doc.sections[0], footer_text)

    add_title_block(doc, content)
    render_summary(doc, content.get("summary", ""))
    render_achievements(doc, content.get("achievements", []))
    render_in_progress(doc, content.get("in_progress", []))
    render_blockers(doc, content.get("blockers", []))
    render_next_week(doc, content.get("next_week", []))

    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description="生成 Word 周报")
    ap.add_argument("--json", required=True, help="周报内容 JSON 路径")
    ap.add_argument("--out", required=True, help="输出 .docx 路径")
    args = ap.parse_args()
    with open(args.json, "r", encoding="utf-8") as f:
        content = json.load(f)
    path = build(content, args.out)
    print("周报已生成:" + path)


if __name__ == "__main__":
    sys.exit(main())
