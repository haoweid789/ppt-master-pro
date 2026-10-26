#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""generate_pptx.py — JSON 驱动的原生可编辑 PPTX 生成器 (ppt-master-pro)

用法:
    python generate_pptx.py slides.json --out output.pptx

JSON 结构见 SKILL.md Phase 3。17 种幻灯片类型 × 12 套主题。
设计规范内置:去 AI 味(无色条/无标题下划线/正文左对齐)、三明治明暗结构、
配色三角色常量、CJK 字体显式设置、原生图表、自动页码与页脚、演讲备注。
生成后自动做溢出/越界粗检并打印报告;坏页跳过不中断,末尾汇总。
"""
import json
import sys
import os
import shutil

# 依赖缺失时必须让 --doctor 仍能跑起来(自检的全部意义就在此),因此延后为软导入
PPTX_IMPORT_ERROR = None
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
except ImportError as e:  # python-pptx 未装:用数值/占位垫片让模块仍可导入
    PPTX_IMPORT_ERROR = e
    Presentation = None

    def Inches(x):
        return int(round(float(x) * 914400))

    def Pt(x):
        return int(round(float(x) * 12700))

    def Emu(x):
        return int(x)

    class _Placeholder:
        """占位命名空间:属性访问一律返回自身,供模块级常量表安全引用。"""

        def __getattr__(self, _name):
            return self

        def __call__(self, *a, **k):
            raise RuntimeError(f"python-pptx 未安装({PPTX_IMPORT_ERROR})")

    RGBColor = PP_ALIGN = MSO_ANCHOR = _Placeholder()
    CategoryChartData = XL_CHART_TYPE = XL_LEGEND_POSITION = XL_LABEL_POSITION = _Placeholder()

# ---------------- 主题(配色三角色常量,逐页只引用不临时取色) ----------------
THEMES = {
    "business_blue": {
        "name": "商业汇报",
        "tags": ["商业", "汇报", "金融", "科技", "制造", "数据", "互联网", "增长", "营收"],
        "bg": "FFFFFF", "bg_dark": "16283F", "primary": "1B4F8A",
        "primary_light": "D9E4F0", "accent": "E8590C",
        "text": "1E293B", "muted": "64748B", "card": "EEF2F7",
        "chart_series": ["1B4F8A", "E8590C", "5B8DC9", "94A3B8", "F59E0B"],
    },
    "academic_white": {
        "name": "学术论文",
        "tags": ["学术", "论文", "答辩", "科研", "教育", "研究", "实验", "开题"],
        "bg": "FFFFFF", "bg_dark": "1F2937", "primary": "334155",
        "primary_light": "E2E8F0", "accent": "B45309",
        "text": "111827", "muted": "6B7280", "card": "F3F4F6",
        "chart_series": ["334155", "B45309", "64748B", "9CA3AF", "D97706"],
    },
    "creative_purple": {
        "name": "创意展示",
        "tags": ["创意", "设计", "品牌", "营销", "文化", "艺术", "广告", "活动"],
        "bg": "FFFFFF", "bg_dark": "2E1065", "primary": "6D28D9",
        "primary_light": "EDE9FE", "accent": "F59E0B",
        "text": "1F2937", "muted": "6B7280", "card": "F5F3FF",
        "chart_series": ["6D28D9", "F59E0B", "A78BFA", "94A3B8", "EF4444"],
    },
    "tech_dark": {
        "name": "技术分享",
        "tags": ["技术", "开发", "代码", "架构", "AI", "算法", "程序", "开源", "工程"],
        "bg": "0F172A", "bg_dark": "020617", "primary": "38BDF8",
        "primary_light": "1E293B", "accent": "34D399",
        "text": "E2E8F0", "muted": "94A3B8", "card": "1E293B",
        "chart_series": ["38BDF8", "34D399", "818CF8", "F472B6", "FBBF24"],
        "dark_body": True,
    },
    "minimal_gray": {
        "name": "通用简约",
        "tags": ["通用", "简约", "咨询", "法律", "总结", "周报", "制度"],
        "bg": "FFFFFF", "bg_dark": "18181B", "primary": "3F3F46",
        "primary_light": "E4E4E7", "accent": "DC2626",
        "text": "18181B", "muted": "71717A", "card": "F4F4F5",
        "chart_series": ["3F3F46", "DC2626", "A1A1AA", "52525B", "F87171"],
    },
    "warm_orange": {
        "name": "消费暖橙(餐饮/零售/生活方式)",
        "tags": ["餐饮", "零售", "消费", "电商", "食品", "门店", "生活方式", "旅游", "外卖"],
        "bg": "FFFFFF", "bg_dark": "431407", "primary": "C2410C",
        "primary_light": "FFEDD5", "accent": "0F766E",
        "text": "1C1917", "muted": "78716C", "card": "FFF7ED",
        "chart_series": ["C2410C", "0F766E", "F59E0B", "A8A29E", "FB923C"],
    },
    "eco_green": {
        "name": "生态绿(健康/环保/农业)",
        "tags": ["健康", "环保", "农业", "绿色", "医疗", "养生", "碳", "可持续", "新能源"],
        "bg": "FFFFFF", "bg_dark": "052E16", "primary": "15803D",
        "primary_light": "DCFCE7", "accent": "CA8A04",
        "text": "14211A", "muted": "6B7F72", "card": "F0FDF4",
        "chart_series": ["15803D", "CA8A04", "4ADE80", "78716C", "0D9488"],
    },
    "finance_green": {
        "name": "金融墨绿",
        "tags": ["金融", "证券", "基金", "财务", "报表", "审计", "银行", "账务", "结算"],
        "bg": "FFFFFF", "bg_dark": "052E16", "primary": "0F766E",
        "primary_light": "CCFBF1", "accent": "C2410C",
        "text": "0F172A", "muted": "64748B", "card": "F0FDFA",
        "chart_series": ["0F766E", "C2410C", "14B8A6", "94A3B8", "F59E0B"],
    },
    "medical_blue": {
        "name": "医疗青",
        "tags": ["医疗", "医院", "护理", "生物", "医药", "临床", "健康", "制药", "诊断"],
        "bg": "FFFFFF", "bg_dark": "0E2A47", "primary": "0369A1",
        "primary_light": "E0F2FE", "accent": "DB2777",
        "text": "0F172A", "muted": "64748B", "card": "F0F9FF",
        "chart_series": ["0369A1", "DB2777", "0EA5E9", "94A3B8", "10B981"],
    },
    "law_burgundy": {
        "name": "法律酒红",
        "tags": ["法律", "法规", "政务", "合规", "政策", "条例", "监管", "裁决", "法务"],
        "bg": "FFFFFF", "bg_dark": "450A0A", "primary": "7F1D1D",
        "primary_light": "FEE2E2", "accent": "166534",
        "text": "1C1917", "muted": "78716C", "card": "FFF1F2",
        "chart_series": ["7F1D1D", "166534", "A16207", "78716C", "B45309"],
    },
    "education_teal": {
        "name": "教育培训青",
        "tags": ["教育", "培训", "校园", "课程", "教学", "学员", "教师", "讲义", "教研"],
        "bg": "FFFFFF", "bg_dark": "134E4A", "primary": "0D9488",
        "primary_light": "CCFBF1", "accent": "F59E0B",
        "text": "134E4A", "muted": "52525B", "card": "F0FDFA",
        "chart_series": ["0D9488", "F59E0B", "14B8A6", "64748B", "0891B2"],
    },
    "luxury_dark": {
        "name": "奢华黑金",
        "tags": ["高端", "奢华", "品牌", "发布会", "奢侈品", "旗舰", "盛典", "峰会"],
        "bg": "0B0B0C", "bg_dark": "000000", "primary": "D4AF37",
        "primary_light": "262626", "accent": "FF6B35",
        "text": "F5F5F4", "muted": "A1A1AA", "card": "1C1917",
        "chart_series": ["D4AF37", "FF6B35", "B87333", "A1A1AA", "888888"],
        "dark_body": True,
    },
}

FONT = "微软雅黑"          # CJK 主字体(Office 自带,无替换风险)
FONT_EN = "Segoe UI"       # 西文辅助
W, H = 13.333, 7.5         # LAYOUT_WIDE 画布
MARGIN = 0.5

SLIDE_W_EMU = Inches(W)
SLIDE_H_EMU = Inches(H)


# ---------------- 基础工具 ----------------
def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _set_run(run, text, theme, size, bold=False, color=None, italic=False, font=FONT):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = _rgb(color or theme["text"])
    # 东亚字体显式声明,防继承默认宋体
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    if ea is None:
        ea = rPr.makeelement(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}ea", {})
        rPr.append(ea)
    ea.set("typeface", font)


def add_textbox(slide, x, y, w, h, lines, theme, size=24, bold=False,
                color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                space_after=6, line_spacing=1.15, font=FONT):
    """lines: str 或 [(text, {size,bold,color,italic}), ...] 列表,每项一段"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [(lines, {})]
    for i, item in enumerate(lines):
        text, style = item if isinstance(item, tuple) else (item, {})
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = style.get("align", align)
        p.space_after = Pt(style.get("space_after", space_after))
        p.line_spacing = style.get("line_spacing", line_spacing)
        run = p.add_run()
        _set_run(run, text, theme,
                 style.get("size", size), style.get("bold", bold),
                 style.get("color", color), style.get("italic", False),
                 style.get("font", font))
    return box


def add_rect(slide, x, y, w, h, fill, line_color=None, line_w=None, shadow=False):
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    shape.adjustments[0] = 0.06
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = _rgb(line_color)
        shape.line.width = Pt(line_w or 1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    if shadow:
        from pptx.enum.shapes import MSO_SHAPE as _S
        el = shape._element.spPr
        # 轻阴影:outer, blur 6pt, offset 2pt, 15% 透明度
        from lxml import etree
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        effect = etree.SubElement(el, "{%s}effectLst" % ns)
        outer = etree.SubElement(effect, "{%s}outerShdw" % ns)
        outer.set("blurRad", "76200")       # 6pt
        outer.set("dist", "25400")          # 2pt
        outer.set("dir", "2700000")         # 45°
        outer.set("rotWithShape", "0")
        clr = etree.SubElement(outer, "{%s}srgbClr" % ns)
        clr.set("val", "000000")
        alpha = etree.SubElement(clr, "{%s}alpha" % ns)
        alpha.set("val", "15000")
    return shape


def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(color)


def add_footer(slide, theme, footer, page_no, total, dark=False):
    color = theme["muted"]
    if footer:
        add_textbox(slide, MARGIN, H - 0.38, 6, 0.3, footer, theme,
                    size=12, color=color)
    if page_no:
        add_textbox(slide, W - MARGIN - 1.2, H - 0.38, 1.2, 0.3,
                    f"{page_no} / {total}", theme, size=12, color=color,
                    align=PP_ALIGN.RIGHT)


def add_title(slide, title, theme, dark=False):
    """内容页标题:左对齐,36pt+,无下划线、无色条(去 AI 味)"""
    color = theme["text"]
    add_textbox(slide, MARGIN, 0.42, W - 2 * MARGIN, 0.9, title, theme,
                size=32, bold=True, color=color)


# ---------------- 17 种幻灯片构建器 ----------------
def s_cover(prs, theme, spec, ctx):
    variant = spec.get("variant", "centered")
    dark = theme.get("dark_body", False)
    if variant == "split":
        # 左文右色块分割(非通栏色条:占半幅的主体分割,属构图而非装饰)
        set_bg(prs.slides[ctx["i"]], theme["bg"])
        add_rect(prs.slides[ctx["i"]], W * 0.62, 0, W * 0.38, H,
                 theme["bg_dark"] if not dark else theme["primary"])
        add_textbox(prs.slides[ctx["i"]], MARGIN + 0.2, H * 0.32, W * 0.5, 2,
                    [(spec["title"], {"size": 48, "bold": True})], theme,
                    color=theme["text"])
        if spec.get("subtitle"):
            add_textbox(prs.slides[ctx["i"]], MARGIN + 0.2, H * 0.32 + 1.6,
                        W * 0.5, 1, spec["subtitle"], theme, size=20,
                        color=theme["muted"])
    else:
        set_bg(prs.slides[ctx["i"]], theme["bg_dark"])
        fg = "FFFFFF"
        sub_fg = theme.get("muted", "94A3B8")
        if theme.get("dark_body"):
            sub_fg = theme["primary"]
        y = H * 0.36 if variant == "centered" else H * 0.30
        align = PP_ALIGN.CENTER if variant == "centered" else PP_ALIGN.LEFT
        x = MARGIN if variant == "left" else MARGIN
        w = W - 2 * MARGIN
        add_textbox(prs.slides[ctx["i"]], x, y, w, 2,
                    [(spec["title"], {"size": 54, "bold": True})], theme,
                    color=fg, align=align)
        if spec.get("subtitle"):
            add_textbox(prs.slides[ctx["i"]], x, y + 1.7, w, 1,
                        spec["subtitle"], theme, size=22, color=sub_fg,
                        align=align)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_section(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg_dark"])
    add_textbox(prs.slides[ctx["i"]], MARGIN, H * 0.40, W - 2 * MARGIN, 1.4,
                spec["title"], theme, size=44, bold=True, color="FFFFFF",
                align=PP_ALIGN.CENTER)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_agenda(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    topics = spec["topics"]
    highlight = spec.get("highlight")
    col_w = (W - 2 * MARGIN - 0.4) / 2
    per_col = (len(topics) + 1) // 2
    for idx, t in enumerate(topics):
        col = 0 if idx < per_col else 1
        row = idx if idx < per_col else idx - per_col
        x = MARGIN + col * (col_w + 0.4)
        y = 1.7 + row * 1.05
        is_hi = (highlight is not None and idx == highlight)
        if is_hi:
            add_rect(prs.slides[ctx["i"]], x, y - 0.12, col_w, 0.85,
                     theme["primary_light"])
        num = add_textbox(prs.slides[ctx["i"]], x + 0.15, y, 0.7, 0.6,
                          f"{idx + 1:02d}", theme, size=26, bold=True,
                          color=theme["accent"] if is_hi else theme["muted"])
        add_textbox(prs.slides[ctx["i"]], x + 0.95, y + 0.06, col_w - 1.1, 0.6,
                    t, theme, size=20, bold=is_hi,
                    color=theme["text"] if is_hi else theme["muted"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_content(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    items = spec["items"]
    columns = spec.get("columns", 1)
    if columns == 2:
        col_w = (W - 2 * MARGIN - 0.4) / 2
        per_col = (len(items) + 1) // 2
        for col in range(2):
            chunk = items[:per_col] if col == 0 else items[per_col:]
            lines = []
            for it in chunk:
                lines.append(("▪ " + it, {"size": 20, "space_after": 14}))
            add_textbox(prs.slides[ctx["i"]], MARGIN + col * (col_w + 0.4),
                        1.8, col_w, H - 2.6, lines, theme)
    else:
        lines = []
        for it in items:
            lines.append(("▪ " + it, {"size": 22, "space_after": 16}))
        add_textbox(prs.slides[ctx["i"]], MARGIN + 0.1, 1.8, W - 2 * MARGIN - 0.2,
                    H - 2.6, lines, theme)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_two_column(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    col_w = (W - 2 * MARGIN - 0.5) / 2
    for col, (heading, items) in enumerate(
            [("left", spec["left"]), ("right", spec["right"])]):
        x = MARGIN + col * (col_w + 0.5)
        head = spec.get(heading + "_title", "方案A" if col == 0 else "方案B")
        add_textbox(prs.slides[ctx["i"]], x, 1.75, col_w, 0.6, head, theme,
                    size=22, bold=True, color=theme["primary"])
        add_rect(prs.slides[ctx["i"]], x, 1.72, 0.0, 0.0, None)  # noop keep z
        lines = [("▪ " + it, {"size": 18, "space_after": 12}) for it in items]
        add_textbox(prs.slides[ctx["i"]], x, 2.45, col_w, H - 3.2, lines, theme)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_table(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    headers, rows = spec["headers"], spec["rows"]
    n_rows, n_cols = len(rows) + 1, len(headers)
    top, tbl_h = 1.9, min(H - 2.6, 0.55 * n_rows)
    shape = prs.slides[ctx["i"]].shapes.add_table(
        n_rows, n_cols, Inches(MARGIN), Inches(top),
        Inches(W - 2 * MARGIN), Inches(tbl_h))
    tbl = shape.table
    for c, htext in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(theme["primary"])
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _set_run(p.add_run(), str(htext), theme, 18, True, "FFFFFF")
    stripe = False
    for r, row in enumerate(rows, start=1):
        fill = theme["card"] if stripe else theme["bg"]
        stripe = not stripe
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(fill)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
            _set_run(p.add_run(), str(val), theme, 16, c == 0, theme["text"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_chart(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    ctype = spec.get("chart_type", "bar")
    data = CategoryChartData()
    data.categories = [str(c) for c in spec["categories"]]
    for s in spec["series"]:
        data.add_series(s["name"], tuple(float(v) for v in s["values"]))
    kind = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE}[ctype]
    x, y, cw, ch = MARGIN + 0.2, 1.85, W - 2 * MARGIN - 0.4, H - 2.7
    if spec.get("source"):
        ch -= 0.35
    gframe = prs.slides[ctx["i"]].shapes.add_chart(
        kind, Inches(x), Inches(y), Inches(cw), Inches(ch), data)
    chart = gframe.chart
    dark = theme.get("dark_body", False)
    txt_color = theme["text"] if not dark else theme["text"]
    chart.has_legend = len(spec["series"]) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(14)
        chart.legend.font.color.rgb = _rgb(theme["muted"])
    plot = chart.plots[0]
    plot.gap_width = 60
    # 系列配色:引用主题 chart_series 常量
    for i, ser in enumerate(chart.series):
        fmt_color = theme["chart_series"][i % len(theme["chart_series"])]
        if ctype == "line":
            ser.format.line.color.rgb = _rgb(fmt_color)
            ser.format.line.width = Pt(2.75)
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = _rgb(fmt_color)
    # 坐标轴文字
    try:
        chart.category_axis.tick_labels.font.size = Pt(14)
        chart.category_axis.tick_labels.font.color.rgb = _rgb(theme["muted"])
        chart.value_axis.tick_labels.font.size = Pt(14)
        chart.value_axis.tick_labels.font.color.rgb = _rgb(theme["muted"])
    except Exception:
        pass
    # 数值标签(bar/pie)
    if ctype in ("bar", "pie"):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(14)
        dl.font.bold = True
        dl.font.color.rgb = _rgb(theme["text"])
        if ctype == "pie":
            dl.show_percentage = True
            dl.show_value = False
    if spec.get("source"):
        add_textbox(prs.slides[ctx["i"]], MARGIN, H - 0.75,
                    W - 2 * MARGIN, 0.3, "数据来源: " + spec["source"], theme,
                    size=12, color=theme["muted"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_timeline(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    ms = spec["milestones"]
    n = len(ms)
    line_y = 4.1  # 内容区垂直居中(标题下方 1.7 至页脚上方 6.9 的中点)
    from pptx.enum.shapes import MSO_SHAPE
    add_rect(prs.slides[ctx["i"]], MARGIN + 0.4, line_y, W - 2 * MARGIN - 0.8,
             0.03, theme["primary"])
    step = (W - 2 * MARGIN - 1.2) / max(n - 1, 1)
    for i, m in enumerate(ms):
        cx = MARGIN + 0.6 + i * step
        dot = prs.slides[ctx["i"]].shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(cx - 0.09), Inches(line_y - 0.075),
            Inches(0.18), Inches(0.18))
        dot.fill.solid()
        dot.fill.fore_color.rgb = _rgb(theme["accent"])
        dot.line.fill.background()
        dot.shadow.inherit = False
        above = (i % 2 == 0)
        box_x = min(max(cx - 1.1, MARGIN), W - MARGIN - 2.2)
        if above:
            add_textbox(prs.slides[ctx["i"]], box_x, line_y - 1.45, 2.2, 0.5,
                        m["label"], theme, size=18, bold=True,
                        color=theme["primary"], align=PP_ALIGN.CENTER)
            add_textbox(prs.slides[ctx["i"]], box_x, line_y - 0.95, 2.2, 0.8,
                        m.get("desc", ""), theme, size=13,
                        color=theme["muted"], align=PP_ALIGN.CENTER)
        else:
            add_textbox(prs.slides[ctx["i"]], box_x, line_y + 0.35, 2.2, 0.5,
                        m["label"], theme, size=18, bold=True,
                        color=theme["primary"], align=PP_ALIGN.CENTER)
            add_textbox(prs.slides[ctx["i"]], box_x, line_y + 0.85, 2.2, 0.8,
                        m.get("desc", ""), theme, size=13,
                        color=theme["muted"], align=PP_ALIGN.CENTER)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_quote(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg_dark"])
    # 装饰性引号:走 Georgia(衬线引号字形正常;雅黑的大号 “ 会看成两道斜杠),
    # 盒子给足高度避免误报溢出
    add_textbox(prs.slides[ctx["i"]], MARGIN + 0.55, 1.3, 1.6, 2.0, "“", theme,
                size=110, bold=False, color=theme["accent"], font="Georgia")
    # 长引文自动降字号,宁可小一号也不压到落款
    quote = spec["quote"]
    qsize = 30 if len(quote) <= 34 else (26 if len(quote) <= 52 else 22)
    add_textbox(prs.slides[ctx["i"]], MARGIN + 0.7, 2.6, W - 2 * MARGIN - 1.4,
                2.4, quote, theme, size=qsize, color="FFFFFF",
                line_spacing=1.3)
    if spec.get("attribution"):
        # 落款跟着引文实际占位走(按字宽估行数),不要硬编码 y 让中间空出一大块
        box_w = W - 2 * MARGIN - 1.4
        cw_in = qsize * 0.85 / 72
        q_lines = max(1, -(-int(len(quote) * cw_in * 100) // int(box_w * 100)))
        attr_y = min(2.6 + q_lines * qsize * 1.3 / 72 + 0.35, 5.6)
        add_textbox(prs.slides[ctx["i"]], MARGIN + 0.7, attr_y,
                    box_w, 0.6, "—— " + spec["attribution"],
                    theme, size=18, color=theme["muted"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_image(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    path = spec["image_path"]
    if not os.path.exists(path):
        raise FileNotFoundError(f"图片不存在: {path}")
    if spec.get("overlay"):
        # 全幅图 + 深色 overlay + 文字(禁止文字直接压图)
        prs.slides[ctx["i"]].shapes.add_picture(
            path, 0, 0, width=SLIDE_W_EMU, height=SLIDE_H_EMU)
        ov = add_rect(prs.slides[ctx["i"]], 0, 0, W, H, "000000")
        ov.fill.fore_color.rgb = _rgb("000000")
        ov.fill.transparency = 0.0
        # python-pptx 无直接透明度 API,用 XML 注入 alpha
        from lxml import etree
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        solidFill = ov._element.spPr.find("{%s}solidFill" % ns)
        clr = solidFill.find("{%s}srgbClr" % ns)
        alpha = etree.SubElement(clr, "{%s}alpha" % ns)
        alpha.set("val", "72000")   # 72% 不透明深色遮罩
        add_textbox(prs.slides[ctx["i"]], MARGIN + 0.4, H * 0.55,
                    W - 2 * MARGIN - 0.8, 1.2, spec["title"], theme,
                    size=36, bold=True, color="FFFFFF")
    else:
        add_title(prs.slides[ctx["i"]], spec["title"], theme)
        # cover 式裁切:按比例铺满内容区,居中裁切(不拉伸)
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        box_w, box_h = W - 2 * MARGIN, H - 2.5
        scale = max(box_w / iw, box_h / ih)
        disp_w, disp_h = iw * scale, ih * scale
        left = MARGIN + (box_w - disp_w) / 2
        top = 1.8 + (box_h - disp_h) / 2
        pic = prs.slides[ctx["i"]].shapes.add_picture(
            path, Inches(left), Inches(top), Inches(disp_w), Inches(disp_h))
        # 裁切到内容区
        if disp_w > box_w:
            frac = (1 - box_w / disp_w) / 2
            pic.crop_left = frac
            pic.crop_right = frac
        if disp_h > box_h:
            frac = (1 - box_h / disp_h) / 2
            pic.crop_top = frac
            pic.crop_bottom = frac
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_summary(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    lines = [("▪ " + p, {"size": 21, "space_after": 12}) for p in spec["points"]]
    add_textbox(prs.slides[ctx["i"]], MARGIN + 0.1, 1.8, W - 2 * MARGIN - 0.2,
                3.0, lines, theme)
    if spec.get("conclusion"):
        band_h = 1.15
        add_rect(prs.slides[ctx["i"]], 0, H - band_h - 0.55, W, band_h,
                 theme["primary_light"])
        add_textbox(prs.slides[ctx["i"]], MARGIN, H - band_h - 0.35,
                    W - 2 * MARGIN, 0.75,
                    [("结论  " + spec["conclusion"],
                      {"size": 22, "bold": True, "color": theme["primary"]})],
                    theme, anchor=MSO_ANCHOR.MIDDLE)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_contact(prs, theme, spec, ctx):
    set_bg(prs.slides[ctx["i"]], theme["bg_dark"])
    add_textbox(prs.slides[ctx["i"]], MARGIN, H * 0.30, W - 2 * MARGIN, 1.4,
                spec.get("title", "谢谢观看"), theme, size=48, bold=True,
                color="FFFFFF", align=PP_ALIGN.CENTER)
    if spec.get("info"):
        add_textbox(prs.slides[ctx["i"]], MARGIN, H * 0.30 + 1.6,
                    W - 2 * MARGIN, 0.8, spec["info"], theme, size=20,
                    color=theme.get("muted", "94A3B8"), align=PP_ALIGN.CENTER)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_matrix(prs, theme, spec, ctx):
    # 对比矩阵:表头+行标签,焦点格(highlight_cells)用浅色底+加粗高亮
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    headers, rows = spec["headers"], spec["rows"]
    focus = {tuple(c) for c in spec.get("highlight_cells", [])}
    shape = prs.slides[ctx["i"]].shapes.add_table(
        len(rows) + 1, len(headers) + 1, Inches(MARGIN), Inches(1.9),
        Inches(W - 2 * MARGIN), Inches(min(H - 2.6, 0.55 * (len(rows) + 1))))
    tbl = shape.table
    # 左上角空格:显式填主色,与表头行连成一条;不填会露出 python-pptx 默认表格样式的品牌蓝
    corner = tbl.cell(0, 0)
    corner.text = ""
    corner.fill.solid()
    corner.fill.fore_color.rgb = _rgb(theme["primary"])
    for c, h in enumerate(headers):
        cell = tbl.cell(0, c + 1)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(theme["primary"])
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _set_run(p.add_run(), str(h), theme, 18, True, "FFFFFF")
    for r, row in enumerate(rows, start=1):
        c0 = tbl.cell(r, 0)
        c0.fill.solid()
        c0.fill.fore_color.rgb = _rgb(theme["primary_light"])
        p0 = c0.text_frame.paragraphs[0]
        _set_run(p0.add_run(), str(row[0]), theme, 17, True, theme["primary"])
        for c, val in enumerate(row[1:], start=1):
            cell = tbl.cell(r, c)
            hl = (r - 1, c - 1) in focus
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(theme["accent"] if hl else theme["card"])
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _set_run(p.add_run(), str(val), theme, 16, hl,
                     "FFFFFF" if hl else theme["text"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_quadrant(prs, theme, spec, ctx):
    # 象限图:坐标轴交叉,点按 x/y(0-100)落位,焦点点用 accent
    from pptx.enum.shapes import MSO_SHAPE
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    area_x, area_y, area_w, area_h = MARGIN + 1.0, 1.9, W - 2 * MARGIN - 2.0, H - 3.3
    cx, cy = area_x + area_w / 2, area_y + area_h / 2   # 坐标原点=区域正中
    # 四象限浅底(只填 1 格权重;行列都用中点,右/下两格才落在画布内)
    for qx, qy in [(area_x, area_y), (cx, area_y), (area_x, cy), (cx, cy)]:
        add_rect(prs.slides[ctx["i"]], qx, qy, area_w / 2, area_h / 2, theme["card"])
    # 轴标签贴各自轴的端头(不压图形):上/下贴竖轴上端与下端,左/右贴横轴两端
    add_textbox(prs.slides[ctx["i"]], cx + 0.1, area_y - 0.45, 3.0, 0.35,
                spec.get("y_axis_top", ""), theme, size=14, color=theme["muted"],
                align=PP_ALIGN.LEFT)
    add_textbox(prs.slides[ctx["i"]], cx + 0.1, area_y + area_h + 0.08, 3.0, 0.35,
                spec.get("y_axis_bottom", ""), theme, size=14, color=theme["muted"],
                align=PP_ALIGN.LEFT)
    add_textbox(prs.slides[ctx["i"]], area_x - 0.92, cy - 0.34, 0.88, 0.3,
                spec.get("x_axis_left", ""), theme, size=14, color=theme["muted"],
                align=PP_ALIGN.RIGHT)
    add_textbox(prs.slides[ctx["i"]], area_x + area_w + 0.06, cy - 0.34, 0.9, 0.3,
                spec.get("x_axis_right", ""), theme, size=14, color=theme["muted"])
    # 轴:横轴贯穿全宽(y=cy),竖轴贯穿全高(x=cx),十字交叉
    axis_h = prs.slides[ctx["i"]].shapes.add_connector(
        1, Inches(area_x), Inches(cy), Inches(area_x + area_w), Inches(cy))
    axis_h.line.color.rgb = _rgb(theme["primary"])
    axis_h.line.width = Pt(2)
    axis_v = prs.slides[ctx["i"]].shapes.add_connector(
        1, Inches(cx), Inches(area_y), Inches(cx), Inches(area_y + area_h))
    axis_v.line.color.rgb = _rgb(theme["primary"])
    axis_v.line.width = Pt(2)
    points = spec["points"]
    hi = set(spec.get("highlight", []))
    for i, pt in enumerate(points):
        px = area_x + pt["x"] / 100 * area_w
        py = area_y + (100 - pt["y"]) / 100 * area_h
        dot = prs.slides[ctx["i"]].shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(px - 0.08), Inches(py - 0.08),
            Inches(0.16), Inches(0.16))
        dot.fill.solid()
        dot.fill.fore_color.rgb = _rgb(theme["accent"] if i in hi else theme["primary"])
        dot.line.fill.background()
        dot.shadow.inherit = False
        lx = px + 0.18 if px + 0.18 + 1.6 < area_x + area_w else px - 1.78
        add_textbox(prs.slides[ctx["i"]], lx, py - 0.12, 1.6, 0.45,
                    [(pt["name"], {"size": 13, "bold": i in hi})]
                    + ([(pt.get("note", ""), {"size": 11, "color": theme["muted"]})]
                       if pt.get("note") else []), theme, line_spacing=1.1)
    if spec.get("source"):
        add_textbox(prs.slides[ctx["i"]], MARGIN, H - 0.75, W - 2 * MARGIN, 0.3,
                    "数据来源: " + spec["source"], theme, size=12,
                    color=theme["muted"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_architecture(prs, theme, spec, ctx):
    # 架构图:分层列表 → 层内横排框 + 层间垂直箭头(全原生连接线/箭头)
    from pptx.enum.shapes import MSO_SHAPE
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    layers = spec["layers"]  # [["层A组件",...], [...], ...]
    top, gap = 1.75, 0.45
    layer_h = (H - top - gap * (len(layers) - 1) - 0.6) / len(layers)
    for li, layer in enumerate(layers):
        ly = top + li * (layer_h + gap)
        n = len(layer)
        box_w = min(3.2, (W - 2 * MARGIN - (n - 1) * 0.25) / n)
        total_w = n * box_w + (n - 1) * 0.25
        x0 = MARGIN + (W - 2 * MARGIN - total_w) / 2
        # 记住每个框的中心,层间箭头按框中心落点(不再用固定偏移,避免箭头插在框边上)
        centers = []
        for idx, b in enumerate(layer):
            bx = x0 + idx * (box_w + 0.25)
            centers.append(bx + box_w / 2)
            box = prs.slides[ctx["i"]].shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(bx), Inches(ly),
                Inches(box_w), Inches(layer_h))
            box.adjustments[0] = 0.1
            box.fill.solid()
            box.fill.fore_color.rgb = _rgb(theme["primary"] if li == 0 else theme["card"])
            box.line.color.rgb = _rgb(theme["primary"] if li == 0 else theme["muted"])
            box.line.width = Pt(1)
            box.shadow.inherit = False
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _set_run(p.add_run(), b, theme, 13, True,
                     "FFFFFF" if li == 0 else theme["text"])
        if li < len(layers) - 1:
            ay_from = ly + layer_h
            ay_to = ly + layer_h + gap
            # 最多画 3 根箭头(组件多时取首/中/末),始终对齐到框中心
            picks = centers if len(centers) <= 3 else [
                centers[0], centers[len(centers) // 2], centers[-1]]
            for fx in picks:
                arrow = prs.slides[ctx["i"]].shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW, Inches(fx - 0.07), Inches(ay_from),
                    Inches(0.14), Inches(ay_to - ay_from))
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = _rgb(theme["muted"])
                arrow.line.fill.background()
                arrow.shadow.inherit = False
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_cards(prs, theme, spec, ctx):
    # 组合卡片行:多张卡片网格(heading + desc + 色块图标区),卡片区分用浅底+阴影,不用边条
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    cards = spec["cards"]
    cols = spec.get("columns", 3)
    rows = (len(cards) + cols - 1) // cols
    top, side = 1.85, MARGIN
    gap_x, gap_y = 0.35, 0.35
    cw = (W - 2 * side - (cols - 1) * gap_x) / cols
    # 卡片高度跟着内容走:上限 3.4in、下限 1.9in,避免一行短文案撑成三个空盒子
    max_ch = (H - top - side - (rows - 1) * gap_y - 0.9) / rows
    desc_lines = 1
    for cd in cards:
        est = int(len(cd.get("desc", "")) * 0.22 / max(cw - 0.5, 1)) + 1
        desc_lines = max(desc_lines, min(est, 8))
    ch = min(3.4, max_ch, max(1.9, 1.05 + desc_lines * 0.26))
    y_off = max(0.0, (H - 0.9 - top - rows * ch - (rows - 1) * gap_y) / 2)
    for i, cd in enumerate(cards):
        r, c = divmod(i, cols)
        x = side + c * (cw + gap_x)
        y = top + y_off + r * (ch + gap_y)
        add_rect(prs.slides[ctx["i"]], x, y, cw, ch, theme["card"], shadow=True)
        sq = add_rect(prs.slides[ctx["i"]], x + 0.18, y + 0.18, 0.42, 0.42,
                      theme["primary"] if i % 2 == 0 else theme["accent"])
        add_textbox(prs.slides[ctx["i"]], x + 0.72, y + 0.22, cw - 0.9, 0.5,
                    cd.get("heading", ""), theme, size=17, bold=True,
                    color=theme["text"], anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(prs.slides[ctx["i"]], x + 0.18, y + 0.75, cw - 0.36, ch - 0.95,
                    cd.get("desc", ""), theme, size=12, color=theme["muted"],
                    line_spacing=1.2)
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


def s_dashboard(prs, theme, spec, ctx):
    # 大数字仪表盘:值 40–52pt + 单位 + 小标签,每行最多 3 个,行间细分隔线
    set_bg(prs.slides[ctx["i"]], theme["bg"])
    add_title(prs.slides[ctx["i"]], spec["title"], theme)
    stats = spec["stats"]
    per_row = 3
    top, side = 1.9, MARGIN
    gap = 0.4
    n_rows = (len(stats) + per_row - 1) // per_row
    # 行高也设上限;剩余空间上下平分,避免数字顶在页顶、下面空半页
    row_h = min(2.5, (H - top - side - 0.7) / n_rows)
    y_off = max(0.0, (H - 0.9 - top - n_rows * row_h) / 2)
    for i, st in enumerate(stats):
        r, c = divmod(i, per_row)
        x = side + c * (W - 2 * side) / per_row
        y = top + y_off + r * row_h
        v = str(st.get("value", ""))
        size = 44 if len(v) <= 6 else 36
        add_textbox(prs.slides[ctx["i"]], x, y, (W - 2 * side) / per_row, 0.9,
                    [(v, {"size": size, "bold": True, "color":
                          theme["accent"] if i % 2 == 1 else theme["primary"]})],
                    theme, align=PP_ALIGN.CENTER)
        if st.get("unit"):
            add_textbox(prs.slides[ctx["i"]], x + ((W - 2 * side) / per_row) / 2 - 1.5,
                        y + 0.72, 3, 0.3, st["unit"], theme, size=14,
                        color=theme["muted"], align=PP_ALIGN.CENTER)
        add_textbox(prs.slides[ctx["i"]], x, y + 1.05, (W - 2 * side) / per_row, 0.4,
                    st.get("label", ""), theme, size=14, color=theme["muted"],
                    align=PP_ALIGN.CENTER)
        # 分隔线只在"下面还有一行"时画,末行之后画线是纯多余的白线
        if r < n_rows - 1:
            ln = prs.slides[ctx["i"]].shapes.add_connector(
                1, Inches(side), Inches(y + row_h + 0.02), Inches(W - side),
                Inches(y + row_h + 0.02))
            ln.line.color.rgb = _rgb(theme["card"])
            ln.line.width = Pt(1)
    if spec.get("source"):
        add_textbox(prs.slides[ctx["i"]], MARGIN, H - 0.75, W - 2 * MARGIN, 0.3,
                    "数据来源: " + spec["source"], theme, size=12,
                    color=theme["muted"])
    if spec.get("notes"):
        prs.slides[ctx["i"]].notes_slide.notes_text_frame.text = spec["notes"]


BUILDERS = {
    "cover": s_cover, "section": s_section, "agenda": s_agenda,
    "content": s_content, "two_column": s_two_column, "table": s_table,
    "chart": s_chart, "timeline": s_timeline, "quote": s_quote,
    "image": s_image, "summary": s_summary, "contact": s_contact,
    "matrix": s_matrix, "quadrant": s_quadrant, "architecture": s_architecture,
    "cards": s_cards, "dashboard": s_dashboard,
}


# ---------------- 删除刚加的空白页(坏页跳过时用) ----------------
def _drop_last_slide(prs):
    """把最后一张幻灯片从 deck 里摘掉。python-pptx 没有公开 API,
    直接操作 sldIdLst + 断掉关系引用,避免留下白页或损坏包结构。"""
    lst = prs.slides._sldIdLst
    items = list(lst)
    if not items:
        return
    sldId = items[-1]
    prs.part.drop_rel(sldId.rId)
    lst.remove(sldId)


# ---------------- 溢出/越界粗检 ----------------
def overflow_check(prs, theme):
    issues = []
    for si, slide in enumerate(prs.slides, start=1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text
            if not txt.strip():
                continue
            # 单字符装饰(引号、序号等)不计入溢出估算,否则每张 quote 页都会误报
            if len(txt.strip()) <= 2:
                continue
            # 越界检测
            if sh.left is None:
                continue
            if sh.left < 0 or sh.top < 0 or \
               sh.left + sh.width > SLIDE_W_EMU + Inches(0.02) or \
               sh.top + sh.height > SLIDE_H_EMU + Inches(0.02):
                issues.append(f"  [越界] 第{si}页: '{txt[:18]}…'")
                continue
            # 粗略溢出:按字号估行高,中文全宽估字宽
            tf = sh.text_frame
            est_h = 0.0
            box_w_in = sh.width / 914400
            for p in tf.paragraphs:
                sizes = [r.font.size.pt for r in p.runs if r.font.size]
                fs = max(sizes) if sizes else 18
                plen = len(p.text)
                cjk = sum(1 for ch in p.text if ord(ch) > 0x2E7F)
                char_w_in = fs * (0.85 if cjk > plen / 2 else 0.52) / 72
                if tf.word_wrap is False:
                    # 不换行文本框:溢出表现为横向出界,行数恒为 1
                    n_lines = 1
                else:
                    n_lines = max(1, -(-int(plen * char_w_in * 100) // int(max(box_w_in, 0.1) * 100)))
                est_h += n_lines * fs * 1.35 / 72
            box_h_in = sh.height / 914400
            if est_h > box_h_in * 1.25:
                issues.append(f"  [可能溢出] 第{si}页: '{txt[:18]}…' "
                              f"估高{est_h:.1f}in > 框高{box_h_in:.1f}in")
    return issues


# ---------------- 风格档案(讯飞 addTheme + Gamma 自定义指令的本地化) ----------------
def load_profile(spec_path):
    """ppt-profile.json 提供跨 deck 默认值(品牌色/字体/页脚/语气档)。
    查找顺序:PPT_PROFILE 环境变量 → spec 同目录 → ~/.ppt-master-pro/;
    spec 里显式写的字段永远优先于档案。"""
    candidates = []
    env = os.environ.get("PPT_PROFILE")
    if env:
        candidates.append(env)
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(spec_path)),
                                   "ppt-profile.json"))
    candidates.append(os.path.join(os.path.expanduser("~"), ".ppt-master-pro",
                                   "ppt-profile.json"))
    for c in candidates:
        if os.path.exists(c):
            try:
                with open(c, encoding="utf-8") as f:
                    prof = json.load(f)
                print(f"[profile] 已加载风格档案 {c}")
                return prof
            except Exception as e:
                print(f"[profile] {c} 解析失败,忽略: {e}")
    return {}


# ---------------- 主题选择(自动匹配:行业词打分,复刻"匹配大纲推荐模板") ----------------
def pick_theme(text):
    """industryHit × 0.6 + 基础分 × 0.4(WPS 逆向报告同构公式);无命中回落 business_blue"""
    best, best_score = "business_blue", 0.0
    for key, th in THEMES.items():
        hits = sum(1 for t in th.get("tags", []) if t in text)
        if hits:
            score = 0.6 * min(hits, 5) / 5 + 0.4
            if score > best_score:
                best, best_score = key, score
    return best


def resolve_theme(spec):
    """custom_theme > style:auto(行业词匹配) > style 指定"""
    if spec.get("custom_theme"):
        ct = {**THEMES["business_blue"], "name": "自定义主题",
              "tags": [], **spec["custom_theme"]}
        THEMES["__custom"] = ct
        return ct, "__custom"
    style = spec.get("style", "business_blue")
    if style == "auto":
        blob = json.dumps(spec, ensure_ascii=False)
        style = pick_theme(blob)
    return THEMES[style], style


# ---------------- 讲稿导出(--script:备注+主张句 → 口语讲稿,说人话双产出) ----------------
def export_script(spec, theme, path):
    lines = [f"# 讲稿 · {spec['slides'][0].get('title', '')}", ""]
    for i, s in enumerate(spec["slides"], 1):
        title = s.get("title") or s.get("quote", "")
        lines.append(f"## 第 {i} 页 · {title}")
        if s.get("notes"):
            lines.append(s["notes"])
        items = s.get("items") or s.get("points") or s.get("topics") or []
        for it in items:
            lines.append(f"- 讲到: {it}")
        lines.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"OK 讲稿已导出 {path}(每页含备注与要点提示,口语化扩写后即可照读)")


# ---------------- 环境自检(--doctor,源自 PPTAgent doctor 命令) ----------------
def doctor():
    print("ppt-master-pro 环境自检")
    ok = True
    # 解释器身份先亮出来:多 Python 并存时,"装到了另一个 python" 是最常见的假缺失
    print(f"  [信息] 解释器 {sys.executable} (Python {sys.version.split()[0]})")
    try:
        import pptx  # noqa: F401
        print(f"  [OK] python-pptx {getattr(pptx, '__version__', '?')}")
    except ImportError:
        print("  [缺失] python-pptx → pip install python-pptx")
        print(f"         注意装给同一个解释器:{sys.executable} -m pip install python-pptx")
        ok = False
    try:
        import defusedxml  # noqa: F401
        print(f"  [OK] defusedxml {getattr(defusedxml, '__version__', '?')}")
    except ImportError:
        print("  [提示] defusedxml 未装:读外部 PPTX(pptx2outline/extract_theme)建议装上")
    try:
        import PIL  # noqa: F401
        print(f"  [OK] Pillow {PIL.__version__}")
    except ImportError:
        print("  [提示] Pillow 未装:仅 image 页 cover 裁切需要 → pip install Pillow")
    soffice = shutil.which("soffice") or (
        r"C:\Program Files\LibreOffice\program\soffice.exe"
        if os.path.exists(r"C:\Program Files\LibreOffice\program\soffice.exe") else None)
    if soffice:
        print(f"  [OK] LibreOffice: {soffice}")
        if not shutil.which("soffice"):
            print("         (不在 PATH,建议注册:setx PATH 补追 ...\\LibreOffice\\program)")
    else:
        print("  [缺失] LibreOffice(视觉 QA 必需,禁用 Keynote/WPS 替代)")
        print("         清华镜像: https://mirrors.tuna.tsinghua.edu.cn/libreoffice/")
        ok = False
    try:
        import shutil as _sh
        if _sh.which("pdftoppm"):
            print("  [OK] pdftoppm (poppler)")
        else:
            print("  [提示] pdftoppm 未装:PDF→PNG 用 → apt/winget 装 poppler")
    except ImportError:
        pass
    print("环境" + ("就绪" if ok else "存在 [缺失] 项:对应能力不可用,补齐后重跑自检"))
    return 0 if ok else 1


# ---------------- 主流程 ----------------
def main():
    import argparse
    ap = argparse.ArgumentParser(description="JSON → 原生可编辑 PPTX")
    ap.add_argument("spec", nargs="?", help="JSON 描述文件(或 .md 大纲,自动调用 md2json)")
    ap.add_argument("--out", help="输出 .pptx 路径(默认取 JSON 的 output 字段)")
    ap.add_argument("--doctor", action="store_true", help="环境自检")
    ap.add_argument("--script", help="同时导出口语讲稿 markdown 到指定路径(说人话双产出)")
    args = ap.parse_args()
    if args.doctor:
        sys.exit(doctor())
    if not args.spec:
        ap.error("需要 spec 参数(JSON 或大纲 .md),或使用 --doctor")
    if PPTX_IMPORT_ERROR is not None:
        print(f"缺少依赖 python-pptx({PPTX_IMPORT_ERROR})。当前解释器:{sys.executable}")
        print(f"  安装:{sys.executable} -m pip install python-pptx")
        print("  或先跑自检:python generate_pptx.py --doctor")
        sys.exit(2)

    if args.spec.lower().endswith(".md"):
        # 大纲 markdown 直译路线(Presenton slides_markdown 思路)
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from md2json import parse_md
        spec = parse_md(args.spec)
    else:
        with open(args.spec, encoding="utf-8") as f:
            spec = json.load(f)

    # 风格档案兜底:spec 显式字段 > 档案(ppt-profile.json)> 内置默认
    prof = load_profile(args.spec)
    for k in ("style", "custom_theme", "footer", "page_numbers", "output"):
        if prof.get(k) is not None:
            spec.setdefault(k, prof[k])

    theme, style_key = resolve_theme(spec)
    out = args.out or spec.get("output", "output.pptx")
    out_dir = os.path.dirname(os.path.abspath(out))
    os.makedirs(out_dir, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_W_EMU
    prs.slide_height = SLIDE_H_EMU

    blank = prs.slide_layouts[6]
    errors = []
    built = []          # [(spec_page, slide)] 只收真正建成的页
    for s in spec["slides"]:
        prs.slides.add_slide(blank)
        idx = len(prs.slides) - 1
        ctx = {"i": idx}
        stype = s.get("type", "content")
        builder = BUILDERS.get(stype)
        if builder is None:
            errors.append(f"  [跳过] 第{idx + 1}页: 未知类型 '{stype}'")
            _drop_last_slide(prs)   # 不留在成品里当白页
            continue
        try:
            builder(prs, theme, s, ctx)
            built.append((s, prs.slides[idx]))
        except Exception as e:
            errors.append(f"  [跳过] 第{idx + 1}页 ({stype}): {e}")
            _drop_last_slide(prs)

    # 页码与页脚(封面不加,quote/contact 深底页加浅色);按"真正建成的页"重新编号
    total = len(built)
    for n, (sp, slide) in enumerate(built):
        if n == 0:
            continue
        dark_page = sp.get("type") in ("section", "quote", "contact")
        add_footer(slide, theme, spec.get("footer"), n + 1, total,
                   dark=dark_page)

    issues = overflow_check(prs, theme)
    prs.save(out)

    print(f"OK 已生成 {out} · 共 {total} 页 · 主题 {theme['name']}"
          + (f"(auto→{style_key})" if spec.get("style") == "auto" and not spec.get("custom_theme") else ""))
    if args.script:
        export_script(spec, theme, args.script)
    if errors:
        print("生成警告:")
        print("\n".join(errors))
    if issues:
        print("溢出粗检(需人工复核或精简文案):")
        print("\n".join(issues))
    if not errors and not issues:
        print("粗检通过:无未知类型、无越界、无明显溢出")
    print("NOT RUN: 视觉 QA(soffice→pdf→png→visual-judge),交付前执行:")
    print(f"  soffice --headless --convert-to pdf --outdir . \"{out}\"")
    print(f"  pdftoppm -png -r 150 \"{os.path.splitext(out)[0]}.pdf\" slide")


if __name__ == "__main__":
    main()
