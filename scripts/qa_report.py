#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""qa_report.py — pptx → 机器可读 QA 报告 JSON(PPTEval 三维的程序化部分)

把"视觉审查-修复循环"的每轮账本变成可 diff 的确定性产物:
每轮生成 QA 报告 → 对比两轮报告 → keep/discard 有实据。

输出 JSON 结构:
{
  "deck": "path", "pages": N, "theme": "name",
  "checks": {
    "content":   {"score": 100, "findings": ["..."]},
    "design":    {"score": 100, "findings": ["..."]},
    "coherence": {"score": 100, "findings": ["..."]}
  },
  "pages": [{"page": 1, "title": "...", "notes_ok": true, "text_len": 23, "pass": true}],
  "pass": true
}

三维检查:
- content: 占位符扫描(xxx/lorem/todo/[insert/Click to add);每页备注存在
- design: 形状越界(画布外);空文本框;文本估计溢出(估算高 vs 框高)
- coherence: 页数;封面存在;结尾页非空
"""
import argparse
import json
import os
import re
import sys

try:
    from pptx import Presentation
except ImportError as _e:  # 依赖缺失给出可执行提示
    sys.exit(f"缺少依赖 python-pptx({_e})。安装:{sys.executable} -m pip install python-pptx")

SLIDE_W = 12192000  # 13.333in
SLIDE_H = 6858000   # 7.5in
EMU_IN = 914400


def scan_content(prs):
    findings = []
    for si, slide in enumerate(prs.slides, start=1):
        txt = slide.shapes[0].text_frame.text if slide.shapes else ""
        full = " ".join(sh.text_frame.text for sh in slide.shapes
                        if sh.has_text_frame)
        for pat in (r"xxx", r"lorem", r"todo", r"\[insert", r"click to add",
                    r"占位", r"待补充"):
            m = re.search(pat, full, re.I)
            if m:
                findings.append(f"第{si}页: 占位符 '{m.group(0)}'")
        if not slide.has_notes_slide or \
            not slide.notes_slide.notes_text_frame.text.strip():
            findings.append(f"第{si}页: 缺演讲备注")
    return findings


def scan_design(prs):
    findings = []
    for si, slide in enumerate(prs.slides, start=1):
        reported = False
        for sh in slide.shapes:
            if sh.left is None or reported:
                continue
            # 越界
            if sh.left < 0 or sh.top < 0 or \
               sh.left + sh.width > SLIDE_W + 18288 or \
               sh.top + sh.height > SLIDE_H + 18288:
                findings.append(f"第{si}页: 形状越界(宽 {sh.width/EMU_IN:.1f}in)")
                reported = True
                continue
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text.strip()
            # 空文本框:只算真文本框(矩形/圆形等自动图形本来就没有文字,不算缺陷)
            if not txt:
                if sh.shape_type is not None and int(sh.shape_type) == 17:
                    findings.append(f"第{si}页: 空文本框")
                    reported = True
                continue
            # 单字符装饰(引号等)不参与溢出估算
            if len(txt) <= 2:
                continue
            # 溢出估算:按字宽推行数,再乘行高(与 generate_pptx 的粗检同一套算法)
            box_w_in = max(sh.width / EMU_IN, 0.1)
            box_h_in = sh.height / EMU_IN
            est = 0.0
            for p in sh.text_frame.paragraphs:
                sizes = [r.font.size.pt for r in p.runs if r.font.size]
                fs = max(sizes) if sizes else 18
                plen = len(p.text)
                if not plen:
                    continue
                cjk = sum(1 for ch in p.text if ord(ch) > 0x2E7F)
                char_w_in = fs * (0.85 if cjk > plen / 2 else 0.52) / 72
                lines = max(1, int(plen * char_w_in / box_w_in) + 1)
                est += lines * fs * 1.35 / 72
            if est > box_h_in * 1.35 and est - box_h_in > 0.35:
                findings.append(
                    f"第{si}页: 文本可能溢出(估 {est:.1f}in > 框 {box_h_in:.1f}in)")
                reported = True
    return findings


def scan_coherence(prs, pages):
    findings = []
    if pages < 2:
        findings.append("页数过少(<2),不构成 deck")
    if not prs.slides[0].shapes or not prs.slides[0].shapes[0].has_text_frame:
        findings.append("封面页无文本")
    last = list(prs.slides)[-1]
    if not any(sh.has_text_frame and sh.text_frame.text.strip()
               for sh in last.shapes):
        findings.append("结尾页无文本")
    return findings


def main():
    ap = argparse.ArgumentParser(description="PPTX → 机器可读 QA 报告")
    ap.add_argument("pptx", help="待检 .pptx")
    ap.add_argument("-o", "--out", help="输出 JSON 路径(默认 <stem>_qa.json)")
    ap.add_argument("--weights", default="0.4,0.4,0.2",
                    help="三维权重 content,design,coherence(默认 0.4,0.4,0.2)")
    args = ap.parse_args()

    prs = Presentation(args.pptx)
    pages = len(prs.slides)

    def score(findings, total_cap=40):
        return max(0, 100 - min(len(findings), total_cap) * (100 / total_cap))

    c = scan_content(prs)
    d = scan_design(prs)
    co = scan_coherence(prs, pages)
    w = [float(x) for x in args.weights.split(",")]
    total_score = (score(c) * w[0] + score(d) * w[1] + score(co) * w[2])

    page_list = []
    for si, slide in enumerate(prs.slides, start=1):
        title = ""
        text_len = 0
        for sh in slide.shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    text_len += len(t)
                    if not title:
                        title = t.split("\n")[0][:30]
        notes_ok = bool(slide.has_notes_slide and
                        slide.notes_slide.notes_text_frame.text.strip())
        issue = []
        if not notes_ok:
            issue.append("缺备注")
        page_list.append({"page": si, "title": title, "notes_ok": notes_ok,
                          "text_len": text_len,
                          "pass": bool(notes_ok and text_len)})

    report = {
        "deck": os.path.abspath(args.pptx),
        "pages": pages,
        "checks": {
            "content": {"score": score(c), "findings": c},
            "design": {"score": score(d), "findings": d},
            "coherence": {"score": score(co), "findings": co},
        },
        "pages": page_list,
        "score": round(total_score, 1),
        "pass": total_score >= 80 and len(c) == 0,
    }
    out = args.out or os.path.splitext(args.pptx)[0] + "_qa.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"OK {out} · {pages}页 · 综合 {report['score']} · "
          f"content {report['checks']['content']['score']} / "
          f"design {report['checks']['design']['score']} / "
          f"coherence {report['checks']['coherence']['score']}")
    if report["pass"]:
        print("PASS 程序化检查全过;视觉审查由 visual-judge 完成(此报告只做账本基线)")
    else:
        print("WARN 需人工复核 findings 后重新生成报告")


if __name__ == "__main__":
    main()
