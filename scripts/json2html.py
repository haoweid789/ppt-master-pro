#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""json2html.py — slides.json → 杂志风单文件网页 PPT(ppt-master-pro 双输出模式)

与 generate_pptx.py 共用同一份 slides.json / outline.md,产出 index.html:
横向翻页、键盘/点击导航、演讲备注按 N 键呼出——源自 ultimate-ppt-master 的
"双输出模式"(可编辑 PPTX ↔ 杂志风网页 deck),分享/手机浏览场景用网页版。

用法:
    python json2html.py slides.json --out index.html
    python generate_pptx.py outline.md --out deck.pptx && python json2html.py outline.md
"""
import json
import sys
import os
import html
import re

THEMES = {
    "business_blue": ("FFFFFF", "16283F", "1B4F8A", "E8590C", "1E293B", "64748B", "EEF2F7"),
    "academic_white": ("FFFFFF", "1F2937", "334155", "B45309", "111827", "6B7280", "F3F4F6"),
    "creative_purple": ("FFFFFF", "2E1065", "6D28D9", "F59E0B", "1F2937", "6B7280", "F5F3FF"),
    "tech_dark": ("0F172A", "020617", "38BDF8", "34D399", "E2E8F0", "94A3B8", "1E293B"),
    "minimal_gray": ("FFFFFF", "18181B", "3F3F46", "DC2626", "18181B", "71717A", "F4F4F5"),
    "warm_orange": ("FFFFFF", "431407", "C2410C", "0F766E", "1C1917", "78716C", "FFF7ED"),
    "eco_green": ("FFFFFF", "052E16", "15803D", "CA8A04", "14211A", "6B7F72", "F0FDF4"),
    "finance_green": ("FFFFFF", "052E16", "0F766E", "C2410C", "0F172A", "64748B", "F0FDFA"),
    "medical_blue": ("FFFFFF", "0E2A47", "0369A1", "DB2777", "0F172A", "64748B", "F0F9FF"),
    "law_burgundy": ("FFFFFF", "450A0A", "7F1D1D", "166534", "1C1917", "78716C", "FFF1F2"),
    "education_teal": ("FFFFFF", "134E4A", "0D9488", "F59E0B", "134E4A", "52525B", "F0FDFA"),
    "luxury_dark": ("0B0B0C", "000000", "D4AF37", "FF6B35", "F5F5F4", "A1A1AA", "1C1917"),
}

E = html.escape


def esc_multiline(t):
    return E(t).replace("\n", "<br>")


def el(tag, cls, inner):
    return f'<{tag} class="{cls}">{inner}</{tag}>'


def bullets(items, cls="list"):
    lis = "".join(f"<li>{E(i)}</li>" for i in items)
    return f'<ul class="{cls}">{lis}</ul>'


def build_table(headers, rows, theme):
    th = "".join(f"<th>{E(str(h))}</th>" for h in headers)
    trs = ""
    for r in rows:
        tds = "".join(f"<td>{E(str(c))}</td>" for c in r)
        trs += f"<tr>{tds}</tr>"
    return f'<table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>'


def build_chart(spec):
    ctype = spec.get("chart_type", "bar")
    cats = [str(c) for c in spec.get("categories", [])]
    series = spec.get("series", [])
    parts = [f'<div class="chart chart-{ctype}">']
    if ctype == "pie":
        vals = series[0]["values"] if series else []
        total = sum(vals) or 1
        acc, stops = 0.0, []
        palette = ["--primary", "--accent", "--muted", "--primary-light", "--card"]
        for i, v in enumerate(vals):
            frac = v / total * 100
            stops.append(f"var({palette[i % len(palette)]}) {acc:.1f}% {acc + frac:.1f}%")
            acc += frac
        parts.append('<div class="pie-wrap">')
        parts.append(f'<div class="pie" style="background:conic-gradient({", ".join(stops)})"></div>')
        legend = "".join(
            f'<span class="lg"><i style="background:var({palette[i % len(palette)]})"></i>'
            f'{E(cats[i])} {v / total * 100:.0f}%</span>'
            for i, v in enumerate(vals))
        parts.append(f'<div class="legend">{legend}</div></div>')
    else:
        n = len(cats)
        if ctype == "bar":
            s0 = series[0] if series else {"values": []}
            mx = max(s0["values"]) if s0["values"] else 1
            bars = ""
            for i, v in enumerate(s0["values"]):
                bars += (f'<div class="bar"><span class="bv">{v:g}</span>'
                         f'<i style="height:{v / mx * 100:.1f}%"></i>'
                         f'<span class="bl">{E(cats[i])}</span></div>')
            parts.append(f'<div class="bars">{bars}</div>')
        else:  # line → SVG
            s0 = series[0] if series else {"values": []}
            vals = s0["values"]
            mx, mn = max(vals) if vals else 1, min(vals) if vals else 0
            rng = (mx - mn) or 1
            w, hgt = 100, 100
            pts = [(i / max(n - 1, 1) * w, hgt - (v - mn) / rng * (hgt - 10) - 5)
                   for i, v in enumerate(vals)]
            poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
            dots = "".join(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="1.6" fill="var(--accent)"/>'
                f'<text x="{x:.1f}" y="{y - 4:.1f}" text-anchor="middle" class="sv">{val:g}</text>'
                for (x, y), val in zip(pts, vals))
            labels = "".join(
                f'<span>{E(c)}</span>' for c in cats)
            parts.append(
                f'<svg viewBox="0 0 {w} {hgt}" preserveAspectRatio="none" class="line-svg">'
                f'<polyline points="{poly}" fill="none" stroke="var(--primary)" '
                f'stroke-width="1.2"/>{dots}</svg>'
                f'<div class="line-labels">{labels}</div>')
    if spec.get("source"):
        parts.append(f'<div class="source">数据来源: {E(spec["source"])}</div>')
    parts.append("</div>")
    return "".join(parts)


def build_slide(spec, idx, total, theme, dark_body):
    t = spec.get("type", "content")
    title = spec.get("title", "")
    notes = spec.get("notes", "")
    dark = t in ("cover", "section", "quote", "contact")
    inner = []
    if t == "cover":
        variant = spec.get("variant", "centered")
        inner.append(f'<div class="cover {variant}">'
                     f'<h1>{E(spec["title"])}</h1>'
                     + (f'<p class="sub">{E(spec["subtitle"])}</p>' if spec.get("subtitle") else "")
                     + "</div>")
    elif t == "section":
        inner.append(f'<h1 class="section-h">{E(title)}</h1>')
    elif t == "agenda":
        topics = spec.get("topics", [])
        hi = spec.get("highlight")
        items = "".join(
            f'<div class="ag{" hi" if hi == i else ""}"><b>{i + 1:02d}</b><span>{E(tp)}</span></div>'
            for i, tp in enumerate(topics))
        inner.append(f'<h2>{E(title)}</h2><div class="agenda">{items}</div>')
    elif t == "content":
        inner.append(f'<h2>{E(title)}</h2>' + bullets(spec.get("items", []), "list big"))
    elif t == "two_column":
        lt, rt = spec.get("left_title", "方案A"), spec.get("right_title", "方案B")
        inner.append(f'<h2>{E(title)}</h2><div class="cols">'
                     f'<div><h3>{E(lt)}</h3>{bullets(spec.get("left", []))}</div>'
                     f'<div><h3>{E(rt)}</h3>{bullets(spec.get("right", []))}</div></div>')
    elif t == "table":
        inner.append(f'<h2>{E(title)}</h2>'
                     + build_table(spec.get("headers", []), spec.get("rows", []), theme))
    elif t == "matrix":
        headers, rows = spec.get("headers", []), spec.get("rows", [])
        focus = {tuple(c) for c in spec.get("highlight_cells", [])}
        th = "".join(f"<th>{E(str(h))}</th>" for h in headers)
        trs = ""
        for r, row in enumerate(rows, start=1):
            tds = (f'<td class="lbl">{E(str(row[0]))}</td>')
            for c, val in enumerate(row[1:], start=1):
                hl = (r - 1, c - 1) in focus
                tds += (f'<td class="{"hl" if hl else ""}">{E(str(val))}</td>')
            trs += f"<tr>{tds}</tr>"
        inner.append(f'<h2>{E(title)}</h2>'
                     f'<table class="matrix"><thead><tr><th class="lbl"></th>{th}</tr></thead>'
                     f'<tbody>{trs}</tbody></table>')
    elif t == "quadrant":
        parts = [f'<h2>{E(title)}</h2><div class="qd">']
        parts.append(f'<div class="qd-axis qx">{E(spec.get("x_axis_left",""))}<span class="mid"></span>{E(spec.get("x_axis_right",""))}</div>')
        parts.append(f'<div class="qd-axis qy-top">{E(spec.get("y_axis_top",""))}</div>')
        parts.append(f'<div class="qd-axis qy-bot">{E(spec.get("y_axis_bottom",""))}</div>')
        parts.append(f'<svg viewBox="0 0 100 100" class="qd-svg" preserveAspectRatio="none">')
        for pt in spec.get("points", []):
            x, y = pt["x"], pt["y"]
            parts.append(f'<circle cx="{x}" cy="{100-y}" r="2" fill="var(--accent)"/>')
            parts.append(f'<text x="{x}" y="{100-y-3}" text-anchor="middle" class="pt-name">{E(pt["name"])}</text>')
        parts.append('</svg>')
        parts.append('</div>')
        inner.append("".join(parts))
    elif t == "architecture":
        parts = [f'<h2>{E(title)}</h2><div class="arch">']
        for li, layer in enumerate(spec.get("layers", [])):
            parts.append(f'<div class="arch-layer{" top" if li==0 else ""}">')
            for b in layer:
                parts.append(f'<span class="arch-box">{E(b)}</span>')
            parts.append("</div>")
        parts.append("</div>")
        inner.append("".join(parts))
    elif t == "cards":
        cards = spec.get("cards", [])
        cols = spec.get("columns", 3)
        cards_html = "".join(
            f'<div class="card"><b>{E(cd.get("heading",""))}</b>'
            f'<p>{E(cd.get("desc",""))}</p></div>'
            for cd in cards)
        inner.append(f'<h2>{E(title)}</h2>'
                     f'<div class="cards" style="grid-template-columns:repeat({cols},1fr)">{cards_html}</div>')
    elif t == "dashboard":
        stats = spec.get("stats", [])
        srows = ""
        for i, st in enumerate(stats):
            srows += (f'<div class="stat"><span class="stat-v" style="color:var(--primary)">'
                      f'{E(str(st.get("value","")))}</span>'
                      f'<span class="stat-l">{E(st.get("label",""))}</span></div>')
        inner.append(f'<h2>{E(title)}</h2><div class="dashboard">{srows}</div>')
    elif t == "chart":
        inner.append(f'<h2>{E(title)}</h2>' + build_chart(spec))
    elif t == "timeline":
        ms = spec.get("milestones", [])
        n = max(len(ms) - 1, 1)
        nodes = ""
        for i, m in enumerate(ms):
            left = 6 + i / n * 88   # 百分比坐标,首尾留 6% 防 label 出界
            side = "up" if i % 2 == 0 else "down"
            nodes += (f'<div class="ms {side}" style="left:{left:.1f}%">'
                      f'<b>{E(m["label"])}</b><span>{E(m.get("desc", ""))}</span></div>')
        inner.append(f'<h2>{E(title)}</h2><div class="tl"><div class="tl-line"></div>{nodes}</div>')
    elif t == "quote":
        q = spec.get("quote", title)
        inner.append(f'<div class="quote-mark">“</div><blockquote>{esc_multiline(q)}</blockquote>'
                     + (f'<p class="attr">—— {E(spec["attribution"])}</p>' if spec.get("attribution") else ""))
    elif t == "image":
        if spec.get("overlay"):
            inner.append(f'<div class="img-full"><img src="{E(spec["image_path"])}">'
                         f'<div class="img-ov"><h2>{E(title)}</h2></div></div>')
        else:
            inner.append(f'<h2>{E(title)}</h2><div class="img-box"><img src="{E(spec["image_path"])}"></div>')
    elif t == "summary":
        concl = (f'<div class="concl"><b>结论</b> {E(spec["conclusion"])}</div>'
                 if spec.get("conclusion") else "")
        inner.append(f'<h2>{E(title)}</h2>' + bullets(spec.get("points", []), "list big") + concl)
    elif t == "contact":
        inner.append(f'<h1 class="section-h">{E(title)}</h1>'
                     + (f'<p class="sub center">{E(spec["info"])}</p>' if spec.get("info") else ""))
    foot = (f'<div class="foot"><span>{E(spec.get("footer", "") or "")}</span>'
            f'<span>{idx + 1} / {total}</span></div>' if idx > 0 else "")
    return (f'<section class="slide {t}{" dark" if dark else ""}" '
            f'data-notes="{E(notes)}">{"".join(inner)}{foot}</section>')


def main():
    import argparse
    ap = argparse.ArgumentParser(description="slides.json → 杂志风网页 PPT")
    ap.add_argument("spec", help="slides.json 或 outline.md")
    ap.add_argument("--out", default=None, help="输出 HTML 路径(默认 <spec名>_deck.html)")
    args = ap.parse_args()

    if args.spec.lower().endswith(".md"):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from md2json import parse_md
        spec = parse_md(args.spec)
    else:
        with open(args.spec, encoding="utf-8") as f:
            spec = json.load(f)

    style = spec.get("style", "business_blue")
    # PPTX 色值无 #,CSS 必须带 # 前缀,否则 var() 解析失败整页透明
    bg, bg_dark, primary, accent, text, muted, card = ("#" + c for c in THEMES[style])
    slides = spec.get("slides", [])
    total = len(slides)
    sections = "".join(
        build_slide(s, i, total, style, style == "tech_dark") for i, s in enumerate(slides))

    out = args.out or (os.path.splitext(args.spec)[0] + "_deck.html")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    doc = f"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(spec.get("output", "deck"))}</title>
<style>
:root{{--bg:{bg};--bg-dark:{bg_dark};--primary:{primary};--accent:{accent};
--text:{text};--muted:{muted};--card:{card}}}
*{{margin:0;padding:0;box-sizing:border-box}}
html{{scroll-snap-type:y mandatory;overflow-y:scroll}}
body{{font-family:"Segoe UI","Microsoft YaHei","PingFang SC",sans-serif;background:var(--bg);color:var(--text)}}
.slide{{min-height:100vh;scroll-snap-align:start;padding:6vh 7vw 9vh;display:flex;
flex-direction:column;justify-content:center;position:relative}}
.slide.dark{{background:var(--bg-dark);color:#fff}}
h1{{font-size:clamp(2rem,4.5vw,4rem);font-weight:800;line-height:1.15}}
h2{{font-size:clamp(1.4rem,2.6vw,2.2rem);font-weight:700;margin-bottom:3.5vh;color:inherit}}
h3{{color:var(--primary);font-size:1.25rem;margin-bottom:1.2vh}}
.cover{{display:flex;flex-direction:column;gap:2vh}}
.cover .sub{{font-size:clamp(1rem,1.8vw,1.5rem);color:var(--muted)}}
.slide.dark .sub{{color:var(--accent)}}
.section-h{{text-align:center}}
.center{{text-align:center}}
ul.list{{list-style:none}}
ul.list li{{padding-left:1.2em;position:relative;margin-bottom:1.6vh;font-size:clamp(1rem,1.6vw,1.35rem);line-height:1.5}}
ul.list li::before{{content:"▪";position:absolute;left:0;color:var(--primary)}}
ul.big li{{font-size:clamp(1.1rem,1.9vw,1.5rem)}}
.cols{{display:grid;grid-template-columns:1fr 1fr;gap:4vw}}
.agenda{{display:grid;grid-template-columns:1fr 1fr;gap:1.6vh 3vw;max-width:70vw}}
.ag{{display:flex;align-items:center;gap:1.2vw;padding:1.4vh 1vw}}
.ag b{{font-size:1.6rem;color:var(--muted)}}
.ag span{{font-size:1.2rem;color:var(--muted)}}
.ag.hi{{background:var(--card);border-radius:.4rem}}
.ag.hi b{{color:var(--accent)}}.ag.hi span{{color:var(--text);font-weight:700}}
table{{border-collapse:collapse;width:100%;font-size:clamp(.9rem,1.4vw,1.15rem)}}
th{{background:var(--primary);color:#fff;padding:1.2vh 1vw;text-align:left}}
td{{padding:1.1vh 1vw;border-bottom:1px solid var(--card)}}
tbody tr:nth-child(even){{background:var(--card)}}
.chart{{flex:1;display:flex;flex-direction:column}}
.bars{{flex:1;display:flex;align-items:flex-end;gap:2.5vw;padding:2vh 1vw 0;min-height:44vh}}
.bar{{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:.6vh;height:44vh}}
.bar i{{width:70%;background:var(--primary);border-radius:.25rem .25rem 0 0;min-height:2px}}
.bar:first-of-type i{{background:var(--accent)}}
.bv{{font-weight:800;color:var(--text)}}
.bl{{color:var(--muted);font-size:.95rem}}
.line-svg{{flex:1;width:100%;min-height:30vh}}
.sv{{font-size:3.2px;font-weight:700;fill:var(--text)}}
.line-labels{{display:flex;justify-content:space-between;color:var(--muted);padding:0 1vw}}
.pie-wrap{{display:flex;align-items:center;gap:4vw}}
.pie{{width:min(34vh,30vw);aspect-ratio:1;border-radius:50%}}
.legend{{display:flex;flex-direction:column;gap:1.4vh;font-size:1.15rem}}
.lg i{{display:inline-block;width:1em;height:1em;margin-right:.5em;border-radius:.15em;vertical-align:-.1em}}
.source{{color:var(--muted);font-size:.85rem;margin-top:2vh}}
.tl{{position:relative;height:44vh}}
.tl-line{{position:absolute;left:0;right:0;top:50%;height:3px;background:var(--primary)}}
.ms{{position:absolute;width:24%}}
.ms b{{color:var(--primary);font-size:1.3rem;display:block}}
.ms span{{color:var(--muted);font-size:.95rem;display:block;margin-top:.4vh}}
.ms.up{{top:50%;transform:translate(-50%,calc(-100% - 5vh));text-align:center}}
.ms.down{{top:50%;transform:translate(-50%,5vh);text-align:center}}
.ms::before{{content:"";position:absolute;left:50%;width:.9rem;height:.9rem;border-radius:50%;
background:var(--accent);transform:translateX(-50%)}}
.ms.up::before{{bottom:-5.8vh}}.ms.down::before{{top:-5.8vh}}
blockquote{{font-size:clamp(1.4rem,2.6vw,2.2rem);line-height:1.5;max-width:70vw}}
.quote-mark{{font-size:6rem;color:var(--accent);line-height:.6;font-family:Georgia,serif}}
.attr{{margin-top:4vh;color:var(--muted);font-size:1.15rem}}
.img-full{{position:absolute;inset:0}}
.img-full img{{width:100%;height:100%;object-fit:cover}}
.img-ov{{position:absolute;inset:0;background:rgba(0,0,0,.55);display:flex;align-items:flex-end;padding:8vh 7vw}}
.img-ov h2{{color:#fff;font-size:clamp(1.6rem,3vw,2.6rem)}}
.img-box{{flex:1;display:flex;align-items:center;justify-content:center}}
.img-box img{{max-width:100%;max-height:56vh;object-fit:cover}}
.concl{{position:absolute;left:0;right:0;bottom:7vh;background:var(--card);
padding:2.2vh 7vw;font-size:clamp(1.05rem,1.7vw,1.4rem);color:var(--primary);font-weight:700}}
.foot{{position:absolute;left:7vw;right:7vw;bottom:2.6vh;display:flex;justify-content:space-between;
color:var(--muted);font-size:.85rem}}
.table.matrix{{width:100%;font-size:clamp(.9rem,1.4vw,1.15rem)}}
.table.matrix th{{background:var(--primary);color:#fff;padding:1.2vh 1vw;text-align:left}}
.table.matrix td{{padding:1.1vh 1vw;border-bottom:1px solid var(--card)}}
.table.matrix td.lbl{{background:var(--card);font-weight:700;color:var(--primary)}}
.table.matrix td.hl{{background:var(--accent);color:#fff;font-weight:700;text-align:center}}
.qd{{position:relative;flex:1;display:flex;min-height:40vh}}
.qd-svg{{flex:1;width:100%;min-height:40vh;border:1px solid var(--card);border-radius:.4rem}}
.pt-name{{font-size:3px;fill:var(--text);font-weight:700;text-anchor:middle}}
.qd-axis{{position:absolute;color:var(--muted);font-size:.85rem;display:flex;gap:2vw}}
.qx{{bottom:.6vh;left:5vw}}.qy-top{{top:1vh;left:2vw}}.qy-bot{{bottom:4vh;left:2vw}}
.arch{{flex:1;display:flex;flex-direction:column;gap:.8vh}}
.arch-layer{{display:flex;gap:1.4vw;justify-content:center;flex-wrap:wrap}}
.arch-box{{border:1.5px solid var(--primary);border-radius:.4rem;padding:1vh 1.2vw;font-size:1.15rem;font-weight:700}}
.arch-layer.top .arch-box{{background:var(--primary);color:#fff}}
.arch-layer.top .arch-box{{border-color:var(--primary)}}
.arch-layer:not(.top) .arch-box{{background:var(--card);color:var(--text);border-color:var(--muted)}}
.cards{{display:grid;gap:1.2vh;grid-template-columns:1fr}}
.card{{background:var(--card);border-radius:.5rem;padding:2.4vh 1.6vw;box-shadow:0 4px 12px rgba(0,0,0,.08)}}
.card b{{font-size:1.25rem;color:var(--primary);display:block;margin-bottom:1vh}}
.card p{{color:var(--muted);font-size:.95rem;line-height:1.5}}
.dashboard{{display:flex;flex-wrap:wrap;gap:2vh 4vw;justify-content:center;align-content:center;flex:1}}
.stat{{display:flex;flex-direction:column;align-items:center;gap:.8vh;min-width:18vw}}
.stat-v{{font-size:clamp(2rem,4vw,3.2rem);font-weight:800}}
.stat-l{{font-size:.95rem;color:var(--muted)}}
.notes-box{{position:fixed;left:0;right:0;bottom:0;background:var(--bg-dark);color:#fff;
padding:2.4vh 7vw;font-size:1.05rem;line-height:1.6;display:none;z-index:9}}
.notes-box.on{{display:block}}
.notes-box b{{color:var(--accent)}}
#tm{{float:right;color:var(--muted);font-variant-numeric:tabular-nums}}
.nav{{position:fixed;right:1.2vw;top:50%;transform:translateY(-50%);display:flex;
flex-direction:column;gap:.7vh;z-index:8}}
.nav i{{width:.5rem;height:.5rem;border-radius:50%;background:var(--muted);opacity:.4;cursor:pointer}}
.nav i.on{{opacity:1;background:var(--accent)}}
@media print{{.nav,.notes-box{{display:none}}.slide{{page-break-after:always}}}}
</style></head><body>
{sections}
<div class="notes-box" id="nb"><b>备注</b> <span id="nt"></span><span id="tm"></span></div>
<div class="nav" id="nav"></div>
<script>
const slides=[...document.querySelectorAll('.slide')],nav=document.getElementById('nav');
slides.forEach((s,i)=>{{const d=document.createElement('i');d.onclick=()=>s.scrollIntoView();nav.appendChild(d)}});
const dots=[...nav.children];
function cur(){{return Math.round(scrollY/innerHeight)}}
function sync(){{const i=cur();dots.forEach((d,j)=>d.classList.toggle('on',j===i));
const n=slides[i]?.dataset.notes;document.getElementById('nb').classList.toggle('on',!!notesOn&&!!n);
if(n)document.getElementById('nt').textContent=n}}
let notesOn=false,t0=null;
function fmt(s){{return String(Math.floor(s/60)).padStart(2,'0')+':'+String(Math.floor(s%60)).padStart(2,'0')}}
setInterval(()=>{{if(notesOn&&t0)document.getElementById('tm').textContent=fmt((Date.now()-t0)/1000)}},500);
addEventListener('scroll',sync,{{passive:true}});
addEventListener('keydown',e=>{{
if(e.key==='ArrowRight'||e.key===' '||e.key==='PageDown')scrollTo({{top:(cur()+1)*innerHeight,behavior:'smooth'}});
else if(e.key==='ArrowLeft'||e.key==='PageUp')scrollTo({{top:(cur()-1)*innerHeight,behavior:'smooth'}});
else if(e.key.toLowerCase()==='n'){{notesOn=!notesOn;if(notesOn&&!t0)t0=Date.now();sync()}}
else if(e.key.toLowerCase()==='r'){{t0=Date.now();document.getElementById('tm').textContent='00:00'}}
else if(e.key==='Home')scrollTo(0,0);else if(e.key==='End')scrollTo(0,document.body.scrollHeight)}});
sync();
</script></body></html>"""
    with open(out, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"OK {out} · {total} 页 · 主题 {style} · 按 N 呼出演讲备注")


if __name__ == "__main__":
    main()
