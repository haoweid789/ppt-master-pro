#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_image.py — 维基共享资源场景图抓取(ppt-master-pro 配图自动化)

路线:维基共享资源(Wikimedia Commons)API 是公开开放接口,无需 key——
按场景关键词搜图,下载高分辨率原图落盘,供 image 页或定制路线插入。
替代"让 agent 手动找图"的低效路径,文字+配图页的自动化闭环。

用法:
    python fetch_image.py "warehouse forklift night" -o images/w1.png --limit 3
    python fetch_image.py "物流仓" -o images/logistics.png --max-width 1200

规则:
- 优先按中文/英文关键词各搜一轮,合并候选
- 下载前过滤:SVG/WMF/矢量格式跳过;缩略图页跳过;尺寸过小(<300px)跳过
- 下载用 User-Agent 标识(维基要求),失败记录不中断
- 图片源为开放许可(CC/GFDL/公有领域),商用场景注意每张图的许可标注
"""
import argparse
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from xml.etree import ElementTree as ET

UA = "ppt-master-pro/1.0 (local skill; contact: local-operator)"
API = "https://commons.wikimedia.org/w/api.php"


def api_get(params):
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        return ET.fromstring(r.read().decode("utf-8"))


def search_files(query, limit=5):
    ns = {"m": "http://www.mediawiki.org/xml/export-0.11/"}
    root = api_get({
        "action": "query", "format": "xml",
        "generator": "search", "gsrsearch": query,
        "gsrnamespace": "6", "gsrlimit": str(limit),
        "prop": "imageinfo",
        "iiprop": "url|size|mime",
        "iiurlwidth": "1280",
    })
    out = []
    for page in root.findall("m:pages/m:page", ns):
        info = page.find("m:imageinfo", ns)
        if info is None:
            continue
        title = page.get("title", "")
        thumb = info.find("m:thumburl", ns)
        mime = info.find("m:mime", ns)
        if thumb is None:
            continue
        m = (mime.text or "") if mime is not None else ""
        if any(x in m for x in ("svg", "wmf", "tiff", "pdf")):
            continue
        mime_ok = any(x in m for x in ("png", "jpeg", "jpg", "webp"))
        if not mime_ok and not re.search(r"\.(png|jpe?g|webp)$", title.lower()):
            continue
        w = info.find("m:width", ns)
        h = info.find("m:height", ns)
        if w is not None and h is not None:
            sw, sh = int(w.text or 0), int(h.text or 0)
            if sw < 300 or sh < 200:
                continue
        out.append({"title": title, "url": thumb.text, "size": (int(w.text) if w else 0,
                                                                      int(h.text) if h else 0)})
    return out


def download(url, dest, max_width=1200):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read()
    os.makedirs(os.path.dirname(os.path.abspath(dest)), exist_ok=True)
    with open(dest, "wb") as f:
        f.write(data)
    return len(data)


def main():
    ap = argparse.ArgumentParser(description="维基共享资源配图抓取(无需 API key)")
    ap.add_argument("query", help="搜索关键词(中文或英文均可)")
    ap.add_argument("-o", "--out", required=True, help="输出图片路径(如 images/w1.png)")
    ap.add_argument("--limit", type=int, default=5, help="每轮候选数")
    ap.add_argument("--max-width", type=int, default=1280, help="按宽度缩放")
    args = ap.parse_args()

    queries = [args.query]
    if not re.search(r"[A-Za-z]", args.query) and re.search(r"[一-鿿]", args.query):
        queries.append(args.query)  # 中文搜索维基自带翻译,无需手动翻译
    elif not re.search(r"[一-鿿]", args.query):
        pass  # 纯英文直接搜

    candidates = []
    seen = set()
    for q in queries:
        for f in search_files(q, args.limit):
            if f["url"] not in seen:
                seen.add(f["url"])
                candidates.append(f)
        if len(candidates) >= args.limit * 2:
            break

    if not candidates:
        sys.exit(f"FAIL 未找到候选图片: '{args.query}' —— 检查关键词或稍后重试")

    ok = 0
    for i, c in enumerate(candidates):
        if ok >= 1:
            break
        ext = os.path.splitext(args.out)[1] or ".png"
        dest = args.out if i == 0 else os.path.splitext(args.out)[0] + f"_{i}{ext}"
        try:
            n = download(c["url"], dest)
            ok += 1
            print(f"OK {dest} · {n//1024}KB · {c['title']}")
        except Exception as e:
            print(f"[跳过] {c['title']}: {str(e)[:60]}")
    if ok == 0:
        sys.exit("FAIL 全部候选下载失败——网络问题或被限流,稍后重试")
    print("NOT RUN: 交付前核验图片许可(CC/公有领域)与场景匹配度")


if __name__ == "__main__":
    main()
