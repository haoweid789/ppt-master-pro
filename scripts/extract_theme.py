#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_theme.py — 参照 PPTX → 主题 JSON(ppt-master-pro 一键美化的数据路径)

复刻 WPS AIPPT 的 extract_ppt_style 思路:分析参照 PPT 的填充色/字色/字体直方图,
映射为 ppt-master-pro 三角色主题(bg/primary/accent/text/muted/card + chart_series),
输出可直接粘进 slides.json 的 "custom_theme" 字段,供 generate_pptx.py 使用。

用法:
    python extract_theme.py reference.pptx                # 打印主题 JSON + 建议
    python extract_theme.py reference.pptx -o theme.json  # 同时落盘

映射启发式(全部按出现频次直方图):
    bg        = 最高频浅色(亮度>230);没有则最高频色
    bg_dark   = 最高频深色(亮度<60),三明治封面底
    primary   = 最高频饱和色(形状填充,饱和度>0.3)
    accent    = 色相距 primary>40° 的最高频饱和色;找不到给警示色
    text      = 字色直方图里最高频深色
    muted     = 最高频中灰(饱和度<0.15,亮度 60–180)
    card/primary_light = primary 与 bg 的浅色混色
    chart_series     = primary、accent + 其余不同色相的高频饱和色,不足补内置序列
结果为启发式估计:交付前用 visual-judge 复核一轮,偏差大的角色手动改。
"""
import json
import colorsys
import argparse
from collections import Counter

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError as _e:  # 依赖缺失给出可执行提示,不抛裸 traceback
    import sys as _sys
    _sys.exit(f"缺少依赖 python-pptx({_e})。安装:{_sys.executable} -m pip install python-pptx")


def _rgb_of(color_format):
    try:
        if color_format and color_format.type is not None:
            return str(color_format.rgb)
    except Exception:
        pass
    return None


def _lum(hexstr):
    r, g, b = (int(hexstr[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _sat_hue(hexstr):
    r, g, b = (int(hexstr[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return s, h * 360


def _mix(a, b, frac_b):
    """a 加 frac_b 权重的 b:用于从 primary 调浅色 tint(frac_b 越大越浅)"""
    ra, ga, ba = (int(a[i:i + 2], 16) for i in (0, 2, 4))
    rb, gb, bb = (int(b[i:i + 2], 16) for i in (0, 2, 4))
    return "%02X%02X%02X" % (round(ra * (1 - frac_b) + rb * frac_b),
                              round(ga * (1 - frac_b) + gb * frac_b),
                              round(ba * (1 - frac_b) + bb * frac_b))


def collect(pptx_path):
    prs = Presentation(pptx_path)
    fills, fonts_colors, font_names = Counter(), Counter(), Counter()

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:  # GROUP
                walk(sh.shapes)
                continue
            try:
                if sh.fill.type is not None and hasattr(sh.fill, "fore_color"):
                    c = _rgb_of(sh.fill.fore_color)
                    if c:
                        # 面积加权:primary 通常铺大面积(色带/图表主色),accent 多为小元素
                        area = (sh.width or 0) * (sh.height or 0) / 914400 ** 2
                        fills[c] += max(1, min(round(area), 40))
            except Exception:
                pass
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        c = _rgb_of(r.font.color)
                        if c:
                            fonts_colors[c] += 1
                        if r.font.name:
                            font_names[r.font.name] += 1
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                c = _rgb_of(r.font.color)
                                if c:
                                    fonts_colors[c] += 1
                                if r.font.name:
                                    font_names[r.font.name] += 1

    for slide in prs.slides:
        try:  # 幻灯片背景不是形状,单独收;权重 ×3 让 bg 在直方图里占主导
            c = _rgb_of(slide.background.fill.fore_color)
            if c:
                fills[c] += 3
        except Exception:
            pass
        walk(slide.shapes)
    return prs, fills, fonts_colors, font_names


def build_theme(fills, fonts_colors, font_names):
    def top(counter, pred):
        for c, _ in counter.most_common():
            if pred(c):
                return c
        return None

    bg = top(fills, lambda c: _lum(c) > 230) or top(fills, lambda c: True) or "FFFFFF"
    bg_dark = top(fills, lambda c: _lum(c) < 60) or "16283F"

    # primary 候选 = 饱和形状填充 + 饱和字色(品牌色通常"铺面积 + 做标题"双出现);
    # 综合分 = 填充面积权重 + 2×字色出现次数
    cand = Counter()
    for c, n in fills.items():
        if _sat_hue(c)[0] > 0.3 and 40 < _lum(c) < 220:
            cand[c] += n
    for c, n in fonts_colors.items():
        if _sat_hue(c)[0] > 0.3 and 40 < _lum(c) < 220:
            cand[c] += 2 * n
    primary = cand.most_common(1)[0][0] if cand else "1B4F8A"
    psat, phue = _sat_hue(primary)

    def cand_score(c):
        return cand.get(c, 0)

    accent = None
    for c, _ in sorted(cand.items(), key=lambda kv: -kv[1]):
        h = _sat_hue(c)[1]
        if min(abs(h - phue), 360 - abs(h - phue)) > 40:
            accent = c
            break
    if accent is None:
        accent = "E8590C" if phue < 20 or phue > 200 else "0F766E"

    text = top(fonts_colors, lambda c: _lum(c) < 90) or top(fills, lambda c: _lum(c) < 90) or "1E293B"
    muted = top(fills, lambda c: _sat_hue(c)[0] < 0.15 and 60 < _lum(c) < 180) or "64748B"

    card = _mix(primary, bg, 0.08)          # 92% bg + 8% primary 的浅底
    primary_light = _mix(primary, bg, 0.16)  # 84% bg + 16% primary

    # chart_series:primary → accent → 其余不同色相的高频饱和色(去重)
    series, used_hues = [primary, accent], [phue]
    for c, _ in fills.most_common(40):
        if c in series:
            continue
        s, h = _sat_hue(c)
        if s < 0.35 or not (40 < _lum(c) < 220):
            continue
        if all(min(abs(h - u), 360 - abs(h - u)) > 45 for u in used_hues):
            series.append(c)
            used_hues.append(h)
        if len(series) >= 5:
            break
    series += ["5B8DC9", "F59E0B", "94A3B8"][:(5 - len(series))]

    main_font = font_names.most_common(1)[0][0] if font_names else "微软雅黑"
    font_ok = any(k in (main_font or "") for k in ("微软雅黑", "雅黑", "苹方", "PingFang", "等线", "宋体"))
    return {
        "name": "提取主题(启发式)",
        "bg": bg, "bg_dark": bg_dark, "primary": primary,
        "primary_light": primary_light, "accent": accent,
        "text": text, "muted": muted, "card": card,
        "chart_series": series[:5],
    }, main_font, font_ok


def main():
    ap = argparse.ArgumentParser(description="参照 PPTX → ppt-master-pro 主题 JSON")
    ap.add_argument("pptx", help="参照 .pptx 文件")
    ap.add_argument("-o", "--out", help="主题 JSON 输出路径(可选)")
    ap.add_argument("--save-profile", action="store_true",
                    help="同时写入 ./ppt-profile.json 风格档案(此后所有 deck 默认使用)")
    args = ap.parse_args()

    _, fills, fonts_colors, font_names = collect(args.pptx)
    if not fills and not fonts_colors:
        print("WARN 未取到任何颜色(文件可能是纯图片型 PPT),返回默认主题")
    theme, main_font, font_ok = build_theme(fills, fonts_colors, font_names)

    result = {"custom_theme": theme}
    js = json.dumps(result, ensure_ascii=False, indent=2)
    print(js)
    print(f"\n主字体: {main_font}" + ("" if font_ok else
          "  ⚠ 观众机器大概率没装这个字体,建议正文换微软雅黑,仅标题保留原字体"))
    print("用法: 粘进 slides.json 顶层字段 \"custom_theme\"(与 style 同级),"
          "generate_pptx.py 会优先使用;")
    print("NOT RUN: 提取结果为启发式估计,交付前渲染一版用 visual-judge 复核配色。")
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"OK 主题已写入 {args.out}")
    if args.save_profile:
        with open("ppt-profile.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print("OK 风格档案 ppt-profile.json 已写入当前目录"
              "(generate_pptx.py 会自动加载;spec 里显式的 style/custom_theme 仍优先)")


if __name__ == "__main__":
    main()
