# ppt-master-pro

端到端专业 PPT 制作技能(v2.6.1):**说人话的内容 × 去 AI 味的设计 × 三输出(可编辑 PPTX / 杂志风网页 deck / 图片版 deck)**。

> 版本历史与全部吸收来源见 [CHANGELOG.md](CHANGELOG.md)。整合自 GitHub 开源技能(ppt-master 系、PPTAgent、Presenton、presentation-ai、academic-pptx-skill、guizang-ppt-skill、codex-ppt、frontend-slides 等)、商业产品逆向(WPS AIPPT/讯飞智文/笔格/歌者/轻竹/Gamma)与方法论项目(karpathy/autoresearch、Slidev)。

## 能力全景

- **三入口**:从零生成 / 编辑式生成(参照 deck 解构)/ 大纲直译(outline.md 一条命令出片)
- **三输出**:可编辑 PPTX / 杂志风单文件网页 deck(横向翻页、N 看备注+计时)/ 图片版 deck(整页生图封装 PPTX+PDF)
- **四法选主题**:风格档案持久化(ppt-profile.json)/ 参照提取(extract_theme.py 一键美化)/ 行业词自动匹配(style:"auto")/ 12 套预设手动点名
- **内容双轨**:叙事轨(杂志风/故事线)vs 事实轨(国际主义/数据驱动)——Phase 0 定轨
- **17 种页面类型**:cover/section/agenda/content/two_column/table/chart/timeline/quote/image/summary/contact + matrix(对比矩阵)/quadrant(象限图)/architecture(架构图)/cards(组合卡片行)/dashboard(大数字仪表盘)
- **说人话四层**:幻灯片主张句 / 演讲备注 / 叙事优势原则 / 口语讲稿(--script)
- **质检**:PPTEval 三维(Content/Design/Coherence)+ 幽灵 deck 测试 + 视觉审查循环(≤2 轮,只修 fail 页)+ 实验账本 + 版本快照
- **配套**:学术模式(引用规范/公式 tex2png/结论页收尾)· 长文档分块大纲 · 环境自检 --doctor · 配图抓取(fetch_image.py)· 结构反解析(pptx2outline.py)· 机器可读 QA 报告(qa_report.py)

## 目录结构

```
ppt-master-pro/
├── SKILL.md                      # 主技能文件(v2.6.1:三入口 × 六阶段 × 三输出)
├── CHANGELOG.md                  # 版本历史与吸收来源
├── scripts/
│   ├── generate_pptx.py          # JSON/MD → 原生可编辑 PPTX(17 类型 × 12 主题 + --script/--doctor)
│   ├── md2json.py                # 大纲 markdown → slides.json 直译
│   ├── json2html.py              # slides.json/MD → 杂志风单文件网页 deck
│   ├── extract_theme.py          # 参照 PPTX → custom_theme JSON(一键美化,--save-profile 存档)
│   ├── build_image_deck.py       # 整页 PNG → 图片版 PPTX/PDF 封装 + 交付三查
│   ├── fetch_image.py            # 维基共享资源配图抓取(无需 API key)
│   ├── pptx2outline.py           # 现有 PPTX → outline.md 结构反解析
│   ├── qa_report.py              # pptx → 机器可读 QA 报告 JSON(PPTEval 确定性基线)
│   └── tex2png.py                # LaTeX 公式 → 透明 PNG(学术模式)
├── references/
│   ├── say-human.md              # 说人话完整规范(文案/备注/叙事/沟通)
│   └── design-techniques.md      # 14 大方向设计技法库
├── templates/
│   ├── outline.md                # 大纲 markdown 语法示例(入口 C 直接改它)
│   ├── slides.json               # JSON 精控骨架(含演讲备注示例)
│   └── slides_all_types.json     # 全 17 种页面类型可运行样例(每页带 notes)
└── README.md
```

## 快速开始

```bash
python scripts/generate_pptx.py --doctor              # 环境自检(先跑这个)
pip install python-pptx                               # 依赖

# 入口 C:大纲直译(改 templates/outline.md 即可出片)
python scripts/generate_pptx.py templates/outline.md --out demo.pptx
python scripts/json2html.py templates/outline.md      # 同一稿 → 网页 deck(N 备注+计时,R 归零)

# 精控:JSON 路线
python scripts/generate_pptx.py templates/slides.json --out demo.pptx
python scripts/generate_pptx.py templates/slides_all_types.json --out demo17.pptx   # 看全 17 类长什么样

# 一键美化:参照 PPT → 主题 → 再生成(extract_ppt_style 复刻)
python scripts/extract_theme.py 参照.pptx --save-profile   # 存为风格档案,此后所有 deck 默认生效
#  或把 theme.json 的 custom_theme 粘进 slides.json;没风格就写 "style":"auto" 按行业词自动选

# 图片版:逐页生图 PNG 存 pages/ 后封装(不假图兜底)
python scripts/build_image_deck.py pages/ --out deck.pptx --pdf deck.pdf --expect <页数>

# 学术公式(Slidev KaTeX 思路,matplotlib mathtext 实现)
python scripts/tex2png.py r"\int_0^1 f(x)\,dx" -o eq.png --color "#1E293B"
```

复杂版式走 SKILL.md Phase 4 定制路线(pptxgenjs / python-pptx 手写);用户给了参照 PPT 先分清模板双轨:"照这个**风格**"=换肤轨(提取主题),"照这个**版式**"=保结构轨(页面级继承)。

## 核心原则(30 秒版)

1. **一页一句主张**:有数字、有对象、有结果;形容词是欠条
2. **配色三角色**:BACKGROUND 60–70% → PRIMARY → ACCENT 5–10%,全篇常量
3. **三明治结构**:封面/封底深底,内容页浅底;一个视觉母题贯穿
4. **去 AI 味**:无色条、无标题下划线、正文左对齐、不堆卡片网格、图表原生
5. **交付前三关**:PPTEval 三维 + 视觉审查循环(≤2 轮)+ 原件保护;明确不做清单见 SKILL.md §8

## 安装为技能

复制整个文件夹到技能目录即可被 Agent 发现:

```bash
# 用户级(所有项目可用)
cp -r ppt-master-pro ~/.agents/skills/
# 或项目级
cp -r ppt-master-pro <project>/.zcode/skills/
```
