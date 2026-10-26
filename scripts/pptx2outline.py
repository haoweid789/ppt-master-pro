#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx2outline.py — 现有 PPTX → outline.md(结构反解析,讯飞 parse-pptx 本地化)

把已有 PPTX 反推为 outline.md(与 md2json.py 语法兼容),实现"拿到别人 deck →
改风格/改内容重出"。extract_theme.py 只解决配色,本脚本解决结构。

用法:
    python pptx2outline.py ref.pptx -o ref_outline.md
    python generate_pptx.py ref_outline.md --out new.pptx   # 改文案后直接重出

推断规则(按页角色):
- 封面/section/quote:文本形状少、字号大或居中 → 原样输出
- 内容页:提取所有文本 → 第一段作标题,其余作要点(- 列表)
- 双栏页:左右两半文本分别输出为 two_column(需要两栏文字量接近)
- 表格页:含表格形状 → table 类型(表头+行)
- 图表页:含图表 → chart 占位说明(数据无法还原,标注 NOT RUN)
- 图片页:含图片 → image 占位说明
"""
import argparse
import os
import re
import sys

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError as _e:  # 依赖缺失给出可执行提示
    sys.exit(f"缺少依赖 python-pptx({_e})。安装:{sys.executable} -m pip install python-pptx")


def slide_texts(slide):
    """提取每页所有文本形状(含表格),返回 [(text, size, left, top)]"""
    out = []
    for sh in slide.shapes:
        if sh.shape_type == 6:  # GROUP
            out.extend(slide_texts(sh))
            continue
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                sizes = [r.font.size.pt for p in sh.text_frame.paragraphs
                         for r in p.runs if r.font.size]
                out.append((t, max(sizes) if sizes else 18,
                            sh.left or 0, sh.top or 0))
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                rowtxt = " | ".join(c.text.strip() for c in row.cells)
                out.append((rowtxt, 14, sh.left or 0, sh.top or 0))
    return out


def detect_role(slide, texts):
    n = len(slide.shapes)
    has_table = any(getattr(sh, "has_table", False) and sh.has_table
                    for sh in slide.shapes)
    has_chart = any(getattr(sh, "has_chart", False) and sh.has_chart
                    for sh in slide.shapes)
    has_pic = any(sh.shape_type == 13 for sh in slide.shapes)
    big = max((t[1] for t in texts), default=0)
    if has_chart:
        return "chart"
    if has_table:
        return "table"
    if n <= 3 and big >= 30:
        return "section"
    if n <= 2 and big >= 36:
        return "cover"
    if has_pic and len(texts) <= 3:
        return "image"
    if texts and texts[0][1] >= 44 and len(texts) == 1:
        return "quote"
    if has_pic:
        return "image"
    return "content"


def extract_title(texts):
    big = max(texts, key=lambda t: t[1])
    return big[0].split("\n")[0][:40]


def extract_items(texts, title_text):
    items = []
    for t, _, _, _ in texts:
        if t == title_text or t.startswith(title_text[:10]):
            continue
        for line in t.split("\n"):
            line = line.strip()
            if line and not line.startswith(("[", "Click", "单击")):
                items.append(line[:60])
    return items[:5]


def detect_two_column(texts):
    """两栏判定:左右文字块量接近且都 >1 行"""
    if len(texts) < 4:
        return None
    # 按 x 分左右
    lefts = [t for t in texts if t[2] < Emu(int(6.5 * 914400))]
    rights = [t for t in texts if t[2] >= Emu(int(6.5 * 914400))]
    if len(lefts) >= 2 and len(rights) >= 2:
        return {"left": [t[0].split("\n")[0] for t in lefts],
                "right": [t[0].split("\n")[0] for t in rights]}
    return None


def main():
    ap = argparse.ArgumentParser(description="PPTX → outline.md 结构反解析")
    ap.add_argument("pptx", help="参考 .pptx")
    ap.add_argument("-o", "--out", help="输出 outline.md(默认 <stem>_outline.md)")
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    out = [f"# 反解析自 {os.path.basename(args.pptx)}",
           "!style: business_blue", "", "# cover | {标题}"]
    # 封面用第一页的标题
    first_texts = slide_texts(prs.slides[0])
    out[-1] = "# cover | " + extract_title(first_texts) if first_texts else "# cover"
    for i, slide in enumerate(list(prs.slides)[1:], start=1):
        texts = slide_texts(slide)
        if not texts:
            continue
        role = detect_role(slide, texts)
        title = extract_title(texts)
        head = f"# {role} | {title}"
        tc = detect_two_column(texts)
        if role == "content" and tc:
            out.append(head)
            out.append(f"# two_column | {title} | 左栏 | 右栏")
            for l in tc["left"]:
                out.append("- " + l)
            out.append("- ==")
            for r in tc["right"]:
                out.append("- " + r)
            continue
        out.append(head)
        if role in ("content", "summary"):
            for it in extract_items(texts, title):
                out.append("- " + it)
        elif role == "quote":
            out.append("> " + title)
        elif role == "section":
            out.append("> " + title)
        elif role == "table":
            out.append("NOT RUN 表格数据可从原文复制或手动确认")
        elif role == "chart":
            out.append("NOT RUN 图表数据无法自动还原,需手动确认")
        elif role == "image":
            out.append("NOT RUN 图片需手动指定路径")
        out.append("")
    out.append("# contact | 谢谢 | 联系方式待补充")

    out_path = args.out or os.path.splitext(args.pptx)[0] + "_outline.md"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    n = len(prs.slides)
    print(f"OK {out_path} · 反解析 {n} 页")
    print("NOT RUN: 表格/图表/图片占位需人工确认后才能重出")
    print("下一步: 编辑 outline.md 文案 → python generate_pptx.py <outline.md> --out new.pptx")


if __name__ == "__main__":
    main()
