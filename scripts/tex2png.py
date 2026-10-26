#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tex2png.py — LaTeX 公式 → 高分辨率透明 PNG(ppt-master-pro 学术模式配套)

依赖 matplotlib mathtext(无需安装 LaTeX),覆盖常用数学符号/上下标/分式/根式/矩阵子集。
复杂公式(宏包、自定义环境)mathtext 不支持——先跑,报错就说明该式超出子集。

用法:
    python tex2png.py "E = mc^2" -o eq1.png
    python tex2png.py r"\\int_0^1 f(x)\\,dx" -o eq2.png --color "#1B4F8A" --size 32

上片:生成的 PNG 用 image 页(overlay=no)或定制路线插入,--color 对齐主题 text 色。
"""
import argparse
import sys

import matplotlib
matplotlib.use("Agg")
from matplotlib import pyplot as plt


def main():
    ap = argparse.ArgumentParser(description="LaTeX 公式 → 透明 PNG(mathtext 子集)")
    ap.add_argument("tex", help="公式源码,不带 $ 定界符")
    ap.add_argument("-o", "--out", default="formula.png", help="输出 PNG 路径")
    ap.add_argument("--size", type=int, default=28, help="等效字号 pt(决定渲染尺寸)")
    ap.add_argument("--color", default="#1E293B", help="文字颜色(带 #,对齐主题 text)")
    ap.add_argument("--dpi", type=int, default=300)
    args = ap.parse_args()

    fig = plt.figure()
    try:
        t = fig.text(0.5, 0.5, f"${args.tex}$", fontsize=args.size,
                     color=args.color, ha="center", va="center")
        fig.canvas.draw()
    except Exception as e:
        sys.exit(f"FAIL mathtext 解析失败(超出子集): {e}\n"
                 f"提示:去掉宏包命令/自定义环境;矩阵用 \\genfrac 或改用图片检索")
    fig.savefig(args.out, dpi=args.dpi, transparent=True,
                bbox_inches="tight", pad_inches=0.06)
    print(f"OK {args.out} · {args.size}pt · dpi{args.dpi}")


if __name__ == "__main__":
    main()
