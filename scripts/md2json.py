#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2json.py — 大纲 Markdown → slides.json(ppt-master-pro 快速路线)

把人类可写的大纲 markdown 转成 generate_pptx.py 的 JSON 输入。
来源:Presenton slides_markdown 思路的本地化——大纲先行、可人工改、跳过二次确认。

语法(逐行解析):
  !style: business_blue        全局设置(style / footer / page_numbers / output)
  # cover | 标题 | 副标题 | left     开新页;extra 段按类型解释:
      cover     → 副标题, 变体(centered/left/split)
      content   → 列数(1 或 2)
      chart     → 图表类型(bar/line/pie)
      two_column→ 左栏标题, 右栏标题
      image     → 图片路径, overlay(yes/no)
      agenda/section/table/summary/quote/timeline/contact → 仅标题
  - 要点                        列表项(agenda=话题 / content·summary=要点 / timeline=「阶段 :: 描述」)
  - ==                          two_column 左右栏分隔线
  | a | b |                     表格行(首行为表头;chart 页首列=类别,其余列=系列)
  > 文本                        按类型解释:quote=署名 · summary=结论 · chart=数据来源 · contact=联系信息
  ! notes: 演讲备注             当前页演讲备注

用法:
  python md2json.py outline.md -o slides.json
  python generate_pptx.py slides.json --out report.pptx
"""
import json
import sys
import argparse


def parse_md(path):
    spec = {"style": "business_blue", "footer": None, "page_numbers": True,
            "output": "output.pptx", "slides": []}
    cur = None
    items, table, left, right, in_right = [], [], [], [], False
    notes, attrib = None, None

    def flush():
        nonlocal cur, items, table, left, right, in_right, notes, attrib
        if cur is None:
            return
        t = cur["type"]
        s = {"type": t, "title": cur["title"]}
        extra = cur.get("extra", [])
        if t == "cover":
            s["subtitle"] = attrib or (extra[0] if extra else None)
            if len(extra) > 1:
                s["variant"] = extra[1]
            elif extra and not attrib:
                s["variant"] = extra[0] if extra[0] in ("centered", "left", "split") else "centered"
        elif t == "content":
            if extra and extra[0] == "2":
                s["columns"] = 2
            s["items"] = items
        elif t == "agenda":
            s["topics"] = items
        elif t == "two_column":
            s["left_title"] = extra[0] if extra else "方案A"
            s["right_title"] = extra[1] if len(extra) > 1 else "方案B"
            if right or in_right:
                s["left"], s["right"] = items, right
            else:  # 无 == 分隔:对半切
                half = (len(items) + 1) // 2
                s["left"], s["right"] = items[:half], items[half:]
        elif t == "table":
            s["headers"] = table[0] if table else []
            s["rows"] = table[1:] if len(table) > 1 else []
        elif t == "matrix":
            if table:
                s["headers"] = table[0][1:]
                s["rows"] = [[r[0]] + r[1:] for r in table[1:]]
        elif t == "cards":
            s["cards"] = [{"heading": it.split("::", 1)[0].strip(),
                           "desc": it.split("::", 1)[1].strip() if "::" in it else ""}
                          for it in items]
        elif t == "dashboard":
            s["stats"] = []
            for it in items:
                if "::" in it:
                    v, label = it.split("::", 1)
                    s["stats"].append({"value": v.strip(), "label": label.strip()})
                else:
                    s["stats"].append({"value": it, "label": ""})
        elif t == "chart":
            s["chart_type"] = extra[0] if extra else "bar"
            if table:
                headers, data_rows = table[0], table[1:]
                s["categories"] = [r[0] for r in data_rows]
                s["series"] = [
                    {"name": headers[i] if i < len(headers) else f"系列{i}",
                     "values": [float(r[i]) if i < len(r) and r[i] else 0
                                for r in data_rows]}
                    for i in range(1, len(headers))]
            if attrib:
                s["source"] = attrib
        elif t == "timeline":
            s["milestones"] = []
            for it in items:
                if "::" in it:
                    label, desc = it.split("::", 1)
                    s["milestones"].append({"label": label.strip(), "desc": desc.strip()})
                else:
                    s["milestones"].append({"label": it, "desc": ""})
        elif t == "quote":
            s["quote"] = cur["title"]
            s["title"] = ""
            if attrib:
                s["attribution"] = attrib
        elif t == "summary":
            s["points"] = items
            if attrib:
                s["conclusion"] = attrib
        elif t == "image":
            if extra:
                s["image_path"] = extra[0]
            if len(extra) > 1 and extra[1].lower() in ("yes", "true", "overlay"):
                s["overlay"] = True
        elif t == "contact":
            if attrib:
                s["info"] = attrib
        if notes:
            s["notes"] = notes
        spec["slides"].append(s)
        cur, items, table, left, right, in_right = None, [], [], [], [], False
        notes, attrib = None, None

    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.rstrip("\n")
            stripped = line.strip()
            if not stripped or stripped.startswith("<!--"):
                continue
            low = stripped.lower()
            if low.startswith("! notes:"):
                notes = stripped[8:].strip()
                continue
            if stripped.startswith("!") and ":" in stripped:
                key, _, val = stripped[1:].partition(":")
                key, val = key.strip().lower(), val.strip()
                if key in ("style", "footer", "output"):
                    spec[key] = val
                elif key == "page_numbers":
                    spec[key] = val.lower() in ("true", "yes", "1", "on")
                continue
            if stripped.startswith("# "):
                flush()
                parts = [p.strip() for p in stripped[2:].split("|")]
                cur = {"type": parts[0].lower(), "title": parts[1] if len(parts) > 1 else "",
                       "extra": [p for p in parts[2:] if p]}
                continue
            if cur is None:
                continue
            if stripped.startswith("|") and stripped.endswith("|"):
                cells = [c.strip() for c in stripped.strip("|").split("|")]
                table.append(cells)
            elif stripped.startswith("- ") or stripped == "-":
                item = stripped[2:].strip()
                if item == "==" and cur["type"] == "two_column":
                    in_right = True
                elif in_right and cur["type"] == "two_column":
                    right.append(item)
                else:
                    items.append(item)
            elif stripped.startswith(">"):
                attrib = stripped[1:].strip()
        flush()
    return spec


def main():
    ap = argparse.ArgumentParser(description="大纲 Markdown → slides.json")
    ap.add_argument("md", help="大纲 markdown 文件")
    ap.add_argument("-o", "--out", default="slides.json", help="输出 JSON 路径")
    args = ap.parse_args()
    spec = parse_md(args.md)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    n = len(spec["slides"])
    no_notes = sum(1 for s in spec["slides"] if not s.get("notes"))
    print(f"OK {args.out} · {n} 页 · 主题 {spec['style']}")
    if no_notes:
        print(f"提醒: {no_notes} 页无演讲备注(说人话规范要求每页备注非空)")


if __name__ == "__main__":
    main()
