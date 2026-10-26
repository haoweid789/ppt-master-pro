---
name: ppt-master-pro
description: 端到端专业 PPT 制作技能（v2.6.1）——三入口（从零生成/编辑式生成/大纲直译）× 六阶段流程 × 三输出（可编辑 PPTX/杂志风网页/图片版）× 四法选主题（档案/提取/自动/手动）× PPTEval 三维评估与视觉审查循环。内置“说人话”表达体系（反 AI 味文案、演讲备注、叙事结构、口语讲稿）与去 AI 味设计规范（配色/字体/版式/图表，17 种页面类型 × 12 套主题 + 学术模式公式支持）。当用户要求制作 PPT、演示文稿、汇报材料、路演 deck、答辩幻灯片，或要求把文档/报告/大纲转成 PPT、美化/编辑现有 PPTX、参照某份 PPT 的风格或版式做新 PPT、写演讲稿与备注时使用。
---

# Skill: ppt-master-pro
# 专业 PPT 制作 × 说人话表达

一句话:交付一份**拿去就能用**的 PPT——每一页有一个具体主张(说人话),整体设计不像 AI 生成的(去 AI 味),文件是原生可编辑的 PPTX(不是图片拼盘)。

**能力全景(v2.6.1)**:三入口(从零/参照/大纲直译)× 三输出(可编辑 PPTX/杂志风网页/图片版)× 四法选主题(档案/提取/自动/手动)× 内容双轨(叙事/事实)× 质量三档 × 说人话四层(文案/备注/叙事/讲稿)× PPTEval 三维 + 实验账本 + 版本快照。版本历史见 `CHANGELOG.md`;**明确不做**见 §8。v2.6.1 为落地校验版:17 种页面类型已逐页渲染复核,修复象限图/矩阵/架构图/仪表盘/卡片/引号页的几何与配色缺陷,补全新类型 JSON 语法文档。

整合来源:hugohe3/ppt-master(SVG→PPTX 原生管线)、gnuhpc/ppt-master-plus(美化/备注/品牌导入)、Offwhite-Del/ppt-master-skill(294 条设计技巧)、editable-ppt-master-skill(双模式生成)、pptx-generator(JSON 驱动)、presentations:pptx(设计规范)、anti-defensive-writing(叙事优势原则)、ui-ux-pro-max(设计检索)、icip-cas/PPTAgent(编辑式生成·PPTEval·视觉审查闭环)、presenton(PPT 排版级模板/图表感知)、allweonedev/presentation-ai(38 主题库/受众风格)、Gabberflast/academic-pptx-skill(学术演示)、22897177sophia-tech/ultimate-ppt-master-skill(输出模式选择器/三输出)、karpathy/autoresearch(实验账本)、Yang9547472171/D-PPT-skill(图片版 deck)、zarazhangrui/frontend-slides(风格小样/单文件 HTML)、Slidev(演讲者视图)、op7418/guizang-ppt-skill(内容风格双轨/排练模式)、kahming/codex-ppt(Markdown-first/角色分离),以及 WPS AIPPT/讯飞智文/笔格PPT/歌者/轻竹/Gamma 六产品逆向。

---

## 0. 三入口模式(先判入口,再进流程)

| 入口 | 判定信号 | 路线 |
|---|---|---|
| A 从零生成 | 只有主题/要点/文档,没有参照 PPT | 全六阶段;素材文档先提炼成页纲 |
| B 编辑式生成(参照 deck) | 用户给了参考 PPT("照这个风格/结构做") | 先做**参照解构**(见 Phase 4 模板继承+页面功能类型学),再按其版式目录填新内容 |
| C 大纲直译 | 用户给的是结构化大纲,或已确认过页纲 | `outline.md → generate_pptx.py` 一条命令出片(Presenton slides_markdown 思路),跳过 Phase 1-2 |

任何入口,Phase 5 质检不减配。

**输出模式选择**(源自 ultimate-ppt-master 风格选择器前置):用户泛泛说"做个 PPT"时,首轮先定输出形态——**PPTX**(默认,可编辑、要交文件)/ **杂志风网页 deck**(单文件 index.html,横向翻页、按 N 看备注,适合分享/手机浏览)/ **图片版 deck**(图像模型整页生图 → PNG → 封装,适合品牌视觉型页面,代价是不可逐对象编辑;D-PPT 路线)。三条路线共用同一份大纲与设计规范;PPTX 与网页可一稿双出(`generate_pptx.py` + `json2html.py`),图片版走独立管线(`build_image_deck.py` 封装,生图用宿主可用接口)。

---

## 0. 两条铁律(先于一切)

1. **内容说人话**:每一页只讲一个具体主张——有数字、有名字、有机制、有"所以呢"。拒绝话题标签、空洞口号、排比堆砌。
2. **设计去 AI 味**:禁止装饰色条/强调条纹、禁止标题下划线、禁止居中正文、禁止连续卡片网格、禁止默认蓝、禁止文字直接压图。完整清单见 §4.6。

**实质先于样式**:深度来自选最锐利的事实并砍掉其余,不来自更多文字。

**原件保护**:用户提供的 PPT 默认只读;交付写新文件 `<stem>_updated.pptx`;仅当用户明确要求原地改时,先在同目录备份 `<stem>_backup.pptx`(绝不放 /tmp)。

---

## 1. 生产流程(六阶段,自动推进不请示)

```text
Phase 0 需求解析 → Phase 1 内容策略 → Phase 2 设计规范
→ Phase 3 内容生成 → Phase 4 视觉实现 → Phase 5 质检交付
```

每阶段产出可判定交付物;阶段间不等待确认;缺参数填占位符继续,最多在最后问一个真正影响下一步的技术问题。

### Phase 0 需求解析(静默完成)

五槽位归一(不展示给用户):

| 槽位 | 内容 |
|---|---|
| Objective | 想要的结果(说服谁/讲清什么/拿到什么决定) |
| Asset | 素材(PDF/DOCX/MD/现有PPT/口述要点/什么都没有) |
| Operation | 从零生成 / 文档转PPT / 美化 / 编辑 / 重建 |
| Deliverable | .pptx + 页数 + 备注 + 讲稿? |
| Verification | 怎样算合格(风格参照/页数/场景) |

场景 × 受众 × 时长 → 页数:每页约 1–1.5 分钟;汇报 10 分钟 ≈ 8–12 页;答辩 15 分钟 ≈ 12–18 页。场景决定基调:工作汇报(结论先行)/学术答辩(证据链)/商业路演(痛点-方案-市场)/教学(循序渐进)。**页数是软目标**(Gamma agent 原则):用户给的页数是期望不是硬约束——内容结构确实需要增减时可以偏离,但必须在 QA 结论里记一句"页数 N→M,原因"。

**内容双轨**(guizang-ppt-skill,Phase 0 定轨,先于 tone):
- **叙事轨(story)**:故事线驱动,偏杂志风/暖色/引语/个人风格——例:品牌故事、个人分享、观点表达。tone 倾向 casual;版式偏引语/时间线/图片主导
- **事实轨(fact)**:数据驱动,偏国际主义/蓝白/图表/方法论——例:产品分析、数据汇报、方法论讲解。tone 倾向 professional;版式偏图表/表格/矩阵
- 判据:用户给的是"想法/故事/感受"还是"数据/分析/结论"。不确定时按事实轨起步,叙事元素做点缀。

**内容三参数**(源自 Presenton/presentation-ai,影响 Phase 1 全部文案):
- **tone**:professional(默认,汇报/答辩)/ casual(内部分享)/ educational(教学,多铺垫多例子)/ sales_pitch(路演,痛点-收益短句)
- **verbosity**:concise(每页 ≤3 行,大字号)/ standard(默认)/ text-heavy(讲义型,可整段,但投影场景慎用)
- **audience**:高管(钱、风险、决策请求)/ 客户(价值、案例)/ 工程师(机制、接口、数据)/ 混合(结论先行+附录深水区)

**检索接地**(事实密集型 deck):关键数字先 WebSearch 核实再上页面,来源写入 source 行;查不到的一律标"示意/est.",不编数。

**学术模式**(会议报告/答辩/基金汇报,源自 academic-pptx-skill 自动启用):
- 交流优先设计:白底、单一无衬线字体、3 色上限、无装饰图标;每结果页**一个 exhibit** + 关键发现直接标注在图上(箭头/高亮框),不在正文复述图上内容
- **引用规范**:借用的每个图/数字加 in-text 引用(图注"图 2,Smith 2024"),deck 末尾加 References 页
- **结论页收尾**:最后一页是 Conclusions(3 条以内,可行动),**不写"谢谢"**——Q&A 环节停在结论页让观众看着结论提问
- **公式支持**(Slidev KaTeX 配套):公式用 `python scripts/tex2png.py "E=mc^2" -o eq.png --color "#<text色>"` 逐式渲染成透明 PNG(matplotlib mathtext,无需 LaTeX),image 页或定制路线插入;行内 ≤2 个符号直接文本写,不贴图

### Phase 1 内容策略(说人话的第一战场)

**叙事结构四选一**:
- SCQA(背景→冲突→问题→答案):适合汇报、路演开场
- 金字塔(结论先行→分论点→论据):适合工作汇报、高管受众
- 时间线/流程:适合项目进展、教学
- 对比(现状 vs 目标 / 我们 vs 竞品):适合方案论证

**素材提炼规则**(源自 anti-defensive-writing 的正面化):
- 每页写一句**主张句**("华南区渠道贡献了 62% 增量"),不是话题标签("渠道分析")
- 数字优先:能写"转化率从 3.1% 提到 5.8%"就不写"大幅提升"
- 围绕最强证据组织全部内容;不占优的维度不设为主战场,换更能体现价值的口径
- 不写成过程汇报("我们先做了A,后来尝试了B")——只保留最终成立的逻辑链
- 翻译成受众的语言:给高管讲钱和风险,给工程师讲机制和接口

**长文档转大纲**(by-file 路线,讯飞智文同思路):输入文档超过 ~30 页时先分块——每块 ≤10 页,逐块摘要(每块提炼 3–5 个候选主张句 + 关键数字),再合并去重成总页纲;合并时做幽灵 deck 测试定去留。禁止把整份长文档直接塞给单次生成。

**页纲先行(强制第一交付物,源自 PPTAgent/Presenton 大纲先行共识)**:先列全部页码 + 每页主张句 + 版式类型 + 页数/时长估算,自查叙事连贯后**先交付页纲**再动手做页;大纲被确认(或用户催促直接做)后转入口 C 直译路线。页纲可写成 `outline.md`(语法见 templates/outline.md)。

**页面脚本**(codex-ppt 精华,页纲的页级细化):页纲是大纲级,生成前每页再写一页脚本——`主张句 | 版式类型 | 视觉焦点(数字/图/表) | 图片需求(描述或路径)`。生成后按脚本逐项核对,claim 是否站住、焦点是否突出、图片是否到位。页面脚本是 final polish 的检查清单,不单独交付,写在页纲的每页备注里。

**幽灵 deck 测试**(源自 academic-pptx-skill,页纲自查的杀手锏):只把每页主张句竖着读一遍——**不看任何图表和正文,主张句序列本身应能讲完整个故事**(问题→方法→发现→含义)。读不通 = 叙事有断点,回炉页纲而不是加内容。

### Phase 2 设计规范(先定后做,全程引用)

**风格小样选型**(frontend-slides 的视觉风格发现,Phase 2 与用户交互的唯一形式):用户说不清审美("好看点""高级感")时,不追问不猜——用现成管线出 **3 个单页小样**(同一页内容 × 3 套主题/版式:一个安全预设、一个大胆方案、一个出格但契合主题的方案),各自渲染成 PNG 让用户挑。小样确认后才铺全 deck,避免整 deck 做完返工。

**配色三角色**:BACKGROUND(占 60–70%)→ PRIMARY(品牌/主题色,标题/结构/图表主色)→ ACCENT(唯一高饱和色,只用在每页最重要的一处,约 5–10%)。加 TEXT 与 MUTED 共五个常量,脚本顶部定义一次,逐页只引用常量,禁止临时取色。

- **取色从主题来**:海洋主题取深海蓝,茶饮品牌取茶褐+奶白;如果这套配色放到毫不相干的 PPT 里也"能用",说明取色不够具体
- **主导而非均分**:一色主导,禁止彩虹配色;需要区分度用主色的深浅变体,不引入新色相
- **三明治结构**:封面/封底深色背景,内容页浅色;或全深色(高级感)
- **一个视觉母题贯穿全篇**:圆角图片框 / 超大数字 / 统一卡片处理,三选一,每页重复。⚠️ 母题绝不能是"色条/侧边条"

**字体**(显式设置到每个 run,永不依赖继承):
- 中文:微软雅黑(Windows)/ 苹方(macOS)为主;思源宋体只用于文化/学术标题;正文禁用楷体/行楷/隶书/艺术字
- 西文按场景配对:科技数据用几何无衬线(Helvetica/Arial/Segoe UI),编辑学术用衬线(Georgia/Source Serif),封面可单独用高对比展示字体
- 最多 2 个字族,层级靠字号+字重,不靠换字体
- CJK 禁用 Light/Thin 字重(投影变糊);正文 Regular,标题 Bold
- 字号(13.33×7.5" 画布):封面标题 44–72pt,内容标题 36pt+,正文 ~24pt,注释/来源 ≥12pt

**版式节奏**:
- 每页一个明确视觉焦点(超大数字/图表/图片/单句金句),其余元素服务于它
- 每页都要有视觉元素(图/表/形状/照片),纯文字页是失败页
- 容器化分组靠对齐+留白+字重,不靠画框:细分隔线 / 全宽行列表 / 1px 描边框 / 浅色通栏带
- 文本页、图片页、数据页交替,控制密度;大留白用内容解决(放大焦点/加真实数据),不用装饰填空

**设计规格锁定**(源自 ultimate-ppt-master 的 design_spec/spec_lock):Phase 2 的产出物是一份 `design_spec.md`——画布尺寸、页数、受众、输出模式(pptx/网页/双出)、配色三角色 hex、字体对、图标与图片策略、视觉母题;**生成前锁定**。此后所有页面只引用锁定值;中途改规格 = 显式 relock(一句话说明改了什么、影响哪些页),禁止逐页漂移。

### Phase 3 内容生成(三条路线)

**路线 C0 · 大纲直译**(最快):写/改 `outline.md` → 直接生成,语法见 `templates/outline.md`:

```bash
python scripts/generate_pptx.py outline.md --out report.pptx   # .md 自动直译 → PPTX
python scripts/generate_pptx.py slides.json --out report.pptx  # JSON 精控 → PPTX
python scripts/json2html.py outline.md                         # 同一稿 → 杂志风网页 deck
python build_image_deck.py 的生图前置:逐页 spec → 图像模型整页生图(宿主可用接口)→ PNG 存 pages/
python scripts/build_image_deck.py pages/ --out deck.pptx --pdf deck.pdf --expect <页数>  # 图片版封装+三查
```

**图片版路线纪律**(D-PPT):生图 prompt 每页一段(版式+内容+风格三要素);**接口失败明确停止,不用占位假图凑数**;生图循环逐页落盘(断点续跑,重跑跳过已完成页);文字准确率是头号质检点,交付前抽页复核。

**路线 C1 · JSON 驱动**:17 种幻灯片类型 + 12 套主题(语法同下,精控每页):

```json
{
  "style": "business_blue",
  "footer": "内部资料",
  "page_numbers": true,
  "output": "report.pptx",
  "slides": [
    {"type": "cover", "title": "标题", "subtitle": "副标题", "variant": "centered|left|split"},
    {"type": "section", "title": "第一部分"},
    {"type": "agenda", "title": "目录", "topics": ["主题1", "主题2"], "highlight": 0},
    {"type": "content", "title": "内容页", "items": ["主张句1", "主张句2"], "columns": 1},
    {"type": "two_column", "title": "对比", "left_title": "现状", "left": ["…"], "right_title": "改造后", "right": ["…"]},
    {"type": "table", "title": "表格", "headers": ["A","B"], "rows": [["x",1,2]]},
    {"type": "chart", "title": "图表", "chart_type": "bar|line|pie",
     "categories": ["Q1","Q2"], "series": [{"name":"S1","values":[10,20]}], "source": "财务系统"},
    {"type": "timeline", "title": "时间线", "milestones": [{"label":"1 月","desc":"立项"}]},
    {"type": "quote", "quote": "内容", "attribution": "作者"},
    {"type": "image", "title": "图片", "image_path": "img.png", "overlay": false},
    {"type": "summary", "title": "总结", "points": ["要点"], "conclusion": "结论"},
    {"type": "contact", "title": "谢谢", "info": "邮箱/电话"}
  ]
}
```

**v2.6 新增 5 类页面语法**(字段名即脚本读取的键,必填项漏了会整页跳过):

```json
[
  {"type": "matrix", "title": "对比矩阵",
   "headers": ["供应商A", "供应商B", "供应商C"],
   "rows": [["时效", "1 天", "1-2 天", "2-3 天"],
            ["成本", "高", "中", "低"]],
   "highlight_cells": [[0, 0], [1, 2]]},

  {"type": "quadrant", "title": "投入产出象限",
   "x_axis_left": "低投入", "x_axis_right": "高投入",
   "y_axis_top": "高回报", "y_axis_bottom": "低回报",
   "points": [{"name": "自动化分拣", "x": 78, "y": 85, "note": "优先做"},
              {"name": "人工复核", "x": 30, "y": 25}],
   "highlight": [0], "source": "内部测算 2026Q3"},

  {"type": "architecture", "title": "系统架构",
   "layers": [["接入层：小程序 / Web / API"],
              ["服务层：订单 / 库存 / 结算"],
              ["数据层：MySQL / Redis / OSS"]]},

  {"type": "cards", "title": "三条改进",
   "columns": 3,
   "cards": [{"heading": "成本", "desc": "单位履约成本降 18%"},
             {"heading": "时效", "desc": "次日达占比 62% → 81%"}]},

  {"type": "dashboard", "title": "关键指标",
   "stats": [{"value": "2.1", "unit": "亿元", "label": "年营收"},
             {"value": "41", "unit": "%", "label": "毛利率"}]}
]
```

| 类型 | 必填 | 可选 | 语义要点 |
|---|---|---|---|
| `matrix` | `title` `headers` `rows` | `highlight_cells` `notes` | 第一列自动当行标签(主色浅底);`rows[0]` 是行标签,**数据列数 = len(headers)**;`highlight_cells` 用 0 基 `[行序, 数据列序]` 指焦点格 |
| `quadrant` | `title` `points[{name,x,y}]` | `x_axis_left/right` `y_axis_top/bottom` `highlight` `source` `notes` | `x/y` 取 0–100 归一化坐标,脚本自己换算落位;`highlight` 是 `points` 的下标数组(用 accent 色);四象限底为单色卡底,不铺四种色 |
| `architecture` | `title` `layers[[组件,…],…]` | `notes` | 数组顺序=自顶向下;第一层用主色实底强调(通常放接入/用户层);层内组件自动居中横排,层间自动画向下箭头,不必手写坐标 |
| `cards` | `title` `cards[{heading,desc}]` | `columns`(默认 3) `notes` | 卡片数不必塞满整行;卡底浅色+阴影区格,不用边条(去 AI 味);描述超 2 行就压字号或改两页 |
| `dashboard` | `title` `stats[{value}]` | `unit` `label` `source` `notes` | 每行固定 3 个,值 ≤6 字符走 44pt、更长降 36pt;`value` 写带单位的字符串会双写单位,数字与 `unit` 分开填 |

五类都是**原生形状/表格**实现(可逐对象编辑),不走贴图;生成后同样进溢出粗检。

```bash
python scripts/generate_pptx.py slides.json --out output.pptx
```

主题(12 套):`business_blue`(商业汇报)/ `academic_white`(学术)/ `creative_purple`(创意)/ `tech_dark`(技术分享)/ `minimal_gray`(通用)/ `warm_orange`(消费·餐饮·生活方式)/ `eco_green`(健康·环保·农业)/ `finance_green`(金融墨绿)/ `medical_blue`(医疗青)/ `law_burgundy`(法律酒红)/ `education_teal`(教育培训青)/ `luxury_dark`(奢华黑金)。取色从主题来,12 套预设覆盖不到就自定三角色常量。

**主题选型四法**(按优先级):
0. **风格档案**(跨 deck 持久化,讯飞 addTheme + Gamma 自定义指令的本地化):品牌色/字体/页脚/主题存一份 `ppt-profile.json`(spec 同目录或 `~/.ppt-master-pro/`),之后每个 deck 自动生效,spec 显式字段仍优先——系列 deck 品牌一致性的正解;`extract_theme.py --save-profile` 提取即存档
1. **提取**(用户给了参照 PPT 想要同款风格,复刻 WPS extract_ppt_style):`python scripts/extract_theme.py 参照.pptx -o theme.json` → 直方图映射出三角色主题 → 把 `custom_theme` 字段粘进 slides.json;字体不适合观众机器时正文换微软雅黑;提取结果交付前渲染复核一轮
2. **自动匹配**(没指定风格):slides.json 里写 `"style": "auto"`,按行业词对主题 tags 打分选预设(餐饮/零售→`warm_orange`,答辩/论文→`academic_white`,AI/架构→`tech_dark`,证券/审计→`finance_green`,医院/临床→`medical_blue`,法规/合规→`law_burgundy`,课程/教研→`education_teal`,发布会/奢侈品→`luxury_dark`…)——每套主题自带 `tags`,匹配靠关键词重合度而非固定映射表
3. **手动指定**:从 12 套预设直接点名

生成时同步写**演讲备注**(add_notes):备注是给"讲的人"的提词,不是幻灯片文字的复读——写开场钩子、强调点、过渡句、可能被问到的问题及答案。备注检查:`notes/total.md` 汇总,每页非空。需要讲稿时加 `--script 讲稿.md`,每页备注+要点导出为口语讲稿(说人话双产出:幻灯片精炼版 + 讲稿口语版)。

### Phase 4 视觉实现(两条技术路线)

**角色分离**(codex-ppt 精华,六阶段天然有角色分离,明确写出来):概念(Phase 1 大纲/页纲)→ 设计(Phase 2 design_spec)→ 配图(Phase 4 视觉素材路由/fetch_image)→ 壳层(Phase 4 渲染实现)。每角色产出独立、接口清晰:概念只管主张句,设计只管配色版式,配图只管素材,壳层只管渲染。角色间不串行阻塞——概念确定后设计/配图可并行推进。

**快速路线**:outline.md / JSON + generate_pptx.py,适合标准结构(汇报/教学/总结)。环境不确定先跑 `python scripts/generate_pptx.py --doctor` 自检依赖。

**定制路线**:pptxgenjs(Node)或 python-pptx(Python)手写,适合复杂版式、模板继承、编辑现有文件。

**排练模式**(guizang-ppt-skill 精华,交付时的可选环节):网页 deck 自带排练能力——N 键呼出备注面板(显示当前页备注+计时器),R 键归零。交付时附一句"排练:浏览器打开 deck.html,按 N 看备注,按 R 归零计时"。PPTX 交付则用备注页在 PowerPoint 演示者视图中排练。

pptxgenjs 关键点:
```javascript
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';               // 13.33×7.5",必须显式设置,默认 10" 会系统性溢出
const W = 13.33, H = 7.5, M = 0.5;
// 颜色永不带 #;透明度用 transparency/opacity 属性,禁 8 位 hex
// bullet 是段落属性,放每个 item 的 options 里;用 { code:"2022", indent:14 } 工厂函数
// 阴影 offset 必须非负;向上投影用 angle:270;选项对象不跨调用复用(工厂函数)
```

图表**永远原生**(addChart / python-pptx chart),继承配色,直接数值标签;堆叠图数据标签只允许 ctr/inBase/inEnd,用 outEnd 会触发 PowerPoint 修复。真正无原生类型的(Sankey/网络图)才降级为图片。

图片三规则:文字永不直接压全幅图(必须加 60–80% 深色 overlay,或改用分栏布局);永不拉伸(比例不符就裁切 cover);灰度打印测试——图上每个字都该能读。

**视觉素材路由**(源自 ppt-master 绘图路由 + Presenton 图库):
- 流程/架构/关系 → **原生形状+连接线**(首选,继承配色)或 Mermaid 导出后矢量插入
- 数据图表 → 原生 chart,永不贴图
- 照片场景图 → 图片检索/生成;检索关键词向"具体场景"取("仓库叉车夜景"不取"物流概念图")
- 图标 → 线性图标库;**禁 emoji 当图标**
- 生成图/AI 图 → 只用于情绪页/封面,数据相关页面禁用(防假数据感)
- **光栅化兜底**(Gamma `pptxBackgroundRasters` 的启示):复杂视觉(密集示意图/渐变艺术背景/3D 结构)OOXML 表达不了或表达后失真时,降级为该区域的高分辨率局部贴图——保真优先于"全是原生对象"的洁癖,但降级要在 QA 结论里记一笔
- **多模型分工**(Gamma 20+ 模型编排的本地对应):正文/备注/大纲用最强文本模型,生图用最强图像模型,各自路由到各自最强的;不在一个模型上硬凑

**模板三轨显式化**(Gamma themeId/templateId + codex-ppt 三层双轨的整合,先问清再动手):
- **风格轨**:整体调性——叙事轨(杂志感)vs 事实轨(国际主义),Phase 0 已定
- **版式轨**:每页布局——从 17 种页面类型中选型(cover/matrix/quadrant/architecture/cards/dashboard 等),按内容双轨选
- **主题轨**:配色字体——换肤(theme:7 预设 + custom_theme + 风格档案)vs 保结构(template:页面级继承)
三轨组合:用户说"照这个风格"是风格轨调整,"照这个版式"是版式轨选型,"照这个配色"是主题轨换肤。三轨可独立变更,互不影响——这就是"改风格/改内容重出"的机制基础。

**模板继承**(保结构轨的具体流程,用户给了 .pptx 模板时强制走;入口 B 的核心):
0. **页面功能类型学**(源自 PPTAgent):先把参照 deck 每页标注为功能类型——cover / TOC / transition / body / data / quote / closing,得到该模板的"版式目录";新内容的每一页从目录里**选型填充**,不从零画
1. 先研究模板:python-pptx 清点每页文本形状(位置/字号/原文长度)、读 theme1.xml 真实色值与字体
2. **预算规则**:`新文本长度 ≤ 原文长度 × 1.1`,这是防溢出的真规则;标签框(短/高字号)只放短语
3. 在副本上改;首选 python-pptx 首个 run 替换(保格式);布局有真占位符用 add_slide,没有就 XML 级克隆整页
4. 高索引先删多余页;克隆页检查每个图片 rId 都能解析到真实媒体
5. 交付 `<stem>_updated.pptx`,原文件字节不动

### Phase 5 质检交付(PPTEval 三维评估 + 视觉审查-修复循环)

**质量档位**(Gamma lite/standard/max 三档的本地化,Phase 0 定档,决定质检预算):

| 档位 | 程序化检查 | 视觉审查 | 适用 |
|---|---|---|---|
| lite | 三维全查 | 跳过(只粗检) | 草稿/内部快速过目 |
| standard(默认) | 三维全查 | visual-judge 抽查 1/3 页 | 日常交付 |
| max | 三维全查 + 幽灵 deck 复测 + 备注 QA 全查 | visual-judge 全页 + 2 轮修复 | 对外交付/重要场合 |

**三维评估**(源自 PPTAgent 的 PPTEval 框架,每维先程序化检查再人工/视觉复核):

| 维度 | 程序化检查 | 通过标准 |
|---|---|---|
| **Content 内容** | grep 占位符(`xxx`/`lorem`/`TODO`/`[insert`);每页有主张句;关键数字有来源行;AI 腔词表扫描(赋能/抓手/闭环/生态) | 零占位、零错字、无 AI 腔、数字可溯源 |
| **Design 设计** | python-pptx 溢出/越界/重叠脚本;去 AI 味清单(§4.1 一票否决项)抽查 XML(无边条/无标题下划线/正文左对齐) | 无溢出重叠、无 AI 味视觉特征 |
| **Coherence 连贯** | 页纲 vs 成品逐页对照;**幽灵 deck 复测**(主张句竖读讲完整故事);叙事结构(SCQA/金字塔)完整;章节过渡页在场 | 结构与页纲一致、翻页有过渡逻辑 |

备注 QA:逐页 notes 非空且不是正文复读。

**视觉审查-修复循环**(源自 PPTAgent reflective generation,最多 2 轮):

```bash
soffice --headless --convert-to pdf --outdir <dir> deck.pptx
pdftoppm -png -r 150 <dir>/deck.pdf <dir>/slide
```

每轮:渲染 → visual-judge 子代理逐页判定(只此一个视觉关口,不自己预筛)→ **只修 fail 的页**(定点修,不全量重做)→ 复渲失败页确认。**单页重绘**(WPS structppt ai_generate_slide 同思路):生成是确定性的,改该页 JSON/MD spec 后整 deck 重跑秒级完成——修单页 = 编辑该页 spec + 重跑 + 复渲该页,不需要页级合并工具。2 轮后仍有 fail:如实报告剩余问题再交付,不无限循环。

**QA 实验账本**(autoresearch keep/discard 思维):每轮审查在 QA 结论里记一行账——`R<N>: <发现的问题> → <修复动作> → <复验结果> → keep/discard`。固定轮次预算(≤2 轮)内账本只增不改;**版本快照**(Gamma snapshot 体系的本地化):每轮修复后的 deck 存 `<stem>_r<N>.pptx`,最终交付版去掉后缀——回滚和对比有实物;账本随 QA 结论一起交付,让"改过什么、为什么"可回溯。禁止无账本的静默重跑。

LibreOffice 缺失时**必须安装**(清华镜像),禁用 Keynote/WPS 替代;装完注册 PATH 并 `--doctor` 验证。

**长 deck 分批**(20 页以上,防上下文超限):每 8–10 页一批生成+粗检,批间核对衔接页(章节过渡),最后整体渲染一次。**断点续跑**(D-PPT):每批完成即落盘,中断后重跑自动跳过已完成批;生图路线按页落盘同理。

**交付物清单**(源自 PPTAgent 交付集):
1. `<name>.pptx` — 成品(新文件,`<stem>_updated.pptx` 规则见铁律);双出模式附 `index.html` 网页版
2. `outline.md` — 页纲(实际生成结构,供下次改版直译)+ `design_spec.md` 规格锁定
3. QA 结论 — 三维评估结果 + 视觉审查轮次与修复记录(一段话,不另开长文档)

原件保护:用户提供的 PPT 默认只读;交付写新文件 `<stem>_updated.pptx`;仅当用户明确要求原地改时,先在同目录备份 `<stem>_backup.pptx`(绝不放 /tmp)。

---

## 2. 说人话体系(让 AI 不说 AI 腔)

### 2.1 幻灯片文案(观众 3 秒 rule)

| AI 腔(禁) | 人话(用) |
|---|---|
| 深度赋能多场景应用生态 | 已接入 3 条产线,单线月省 11 万 |
| 构建了全链路数字化闭环 | 从下单到出货,系统全程盯 |
| 多维度、全方位提升用户体验 | 客诉响应从 48h 压到 4h |
| 抢占行业新赛道,引领变革 | 明年做到区域前三,份额 15% |
| 依托核心算法引擎驱动创新 | 换了匹配算法,点击率翻倍 |

规则:
- **有数字、有对象、有结果**;没有数字时给具体事实(名字/地点/机制/对比)
- 形容词是欠条:"显著提升"= 没量;"行业领先"= 没证据
- 一页一句主张,其余是支撑;砍掉所有"正确的废话"
- 口语化有度:汇报可以口语,答辩保持术语准确

### 2.2 演讲备注(讲的人的话)

备注三段式:①这页怎么开口(一句钩子/承接上一页) ②哪个数字/词要停顿强调 ③预判提问+应答。禁止把幻灯片文字复制进备注。

### 2.3 叙事说人话(优势原则)

- 论文/汇报是发布会,不是工作日志:只讲最终成立的逻辑,不讲"我们试过什么"
- 优势明确说出来,不等听众自己从表格里发现
- 不主动递刀:结论段不突然自我否定;局部缺点不上升为整体缺陷
- 控制主张范围:只提证据撑得住的话,说服力来自主张与证据一致

### 2.4 与用户沟通说人话

- 第一句就是结果/交付物,不绕"好的,我来帮您…"
- 汇报进度给事实(页数/文件路径/检查结果),不给过程独白
- 出问题时:一句说明 + 已走的替代路径,不铺陈困难

---

## 3. 版式库(按内容选型)

| 内容 | 版式 |
|---|---|
| 开场/封底 | 大字标题布局(深底) |
| 流程/工作流 | 编号步骤链(chevron/箭头/时间线) |
| 对比 | 左右分栏 / 对比表 |
| 数据/指标 | 图表为主 + 超大数字(60–72pt) |
| 关键洞察 | 单句强调布局 |
| 团队/人物 | 并排卡片行(一行,非网格) |
| 架构/关系 | 概念示意图(框+连接线) |
| 引用/金句 | 全屏引号布局 |

版式按页轮换:文本页→图示页→数据页→金句页;禁止连续同版式。

---

## 4. 去AI味设计清单(违反即返工)

### 4.1 绝对禁止(一票否决)
1. 标题下加装饰下划线
2. 装饰色条/强调条纹:通栏 header/footer 色带、竖向侧边条、卡片单侧色边、连续多页同款边条
3. 正文居中(只有标题/金句可居中)
4. 文字直接压全幅图无 overlay
5. 拉伸变形图片
6. 幻灯片内容里用 emoji
7. 带 `#` 的 hex(文件损坏)/ 8 位 hex 当透明度
8. 连续 2 页以上卡片网格;网格占比超过 1/5 内容页

### 4.2 强烈避免
- 默认蓝当主色;奶油色/米色默认背景(用纯白 FFFFFF 或品牌色)
- 每页同版式;一页塞多个重点;字号层级拉不开(标题需 36pt+ 且与正文对比明显)
- 粗 bullet 默认样式(用 `{code:"2022", indent:14}`,缩进 10–16pt)
- 渐变、3D、阴影滥用(卡片区分用浅底色或轻阴影,不用边条)
- 大留白用装饰填空(用内容填:放大焦点/真实数据/金句)

### 4.3 必做
- 13.33×7.5" 画布显式声明;最小边距 0.5";块间距统一 0.3 或 0.5
- 每个关键数字/图表配来源行(10–12pt muted:"Source: IDC, 2025");估算值标注"示意/est."
- CJK 文本框比西文宽 ~15%;行短,防中段折行

---

## 5. 依赖与命令

```bash
python scripts/generate_pptx.py --doctor    # 环境自检(依赖不确定时先跑)
pip install python-pptx defusedxml          # Python 路线
npm install -g pptxgenjs                     # Node 路线
# LibreOffice(渲染质检,缺失必须装,禁替代品):
#   win: winget install TheDocumentFoundation.LibreOffice(或清华镜像 msi)
#   mac: brew install --cask libreoffice
#   linux(deb): 清华镜像 tar.gz → dpkg -i
soffice --headless --convert-to pdf --outdir <dir> deck.pptx   # PPTX→PDF
pdftoppm -png -r 150 <dir>/deck.pdf <dir>/slide                # PDF→PNG
# 网页 deck → PDF:无头浏览器逐页 scrollTo + 截图 → build_image_deck.py --pdf 封装
#   (或直接浏览器打印,HTML 内置 print CSS)
```

字体注意:预览用 LibreOffice 渲染,没有的字体会被替换导致宽度失真——优先用 Office 自带字体(微软雅黑/等线/苹方),替换风险高的字体留 10% 宽度余量。

---

## 6. 交付引用规范

最终回复中,每个交付文件用一次 `::zcode-file-citation{path="…" purpose="output"}` 内联引用;问答/未改动用 `purpose="source"` 并带 slide_number。不在结尾重复罗列文件名/路径/链接。

---

## 7. 扩展参考(按需加载)

| 文件 | 内容 | 何时读 |
|---|---|---|
| `references/say-human.md` | 说人话完整规范:文案改写对照库、备注模板、叙事原则全文 | 写文案/备注卡壳时 |
| `references/design-techniques.md` | 14 大方向设计技巧(版式9式/色彩70-25-5/图表/图片/动画/逻辑模型) | 定设计规范时 |
| `scripts/generate_pptx.py` | JSON/MD→PPTX 生成器(17 类型 × 12 主题 + --doctor 自检) | 快速路线执行时 |
| `scripts/md2json.py` | 大纲 markdown → slides.json 直译(可单独调用) | 入口 C / 页纲转译时 |
| `scripts/json2html.py` | slides.json/MD → 杂志风单文件网页 deck(横向翻页,N 呼出备注) | 双输出模式 / 分享场景时 |
| `scripts/extract_theme.py` | 参照 PPTX → custom_theme JSON(一键美化的数据路径) | 用户给了参照 PPT 想要同款风格时 |
| `scripts/build_image_deck.py` | 整页 PNG → 图片版 PPTX/PDF 封装 + 交付三查 | 图片版输出模式 / 品牌视觉型页面时 |
| `scripts/tex2png.py` | LaTeX 公式 → 透明 PNG(mathtext 子集) | 学术模式 / 公式多的 deck |
| `scripts/fetch_image.py` | 维基共享资源配图抓取(无需 key) | 文字+配图页自动化 |
| `scripts/pptx2outline.py` | 现有 PPTX → outline.md 结构反解析 | "拿到别人 deck → 改风格/内容重出" |
| `scripts/qa_report.py` | pptx → 机器可读 QA 报告 JSON(PPTEval 三维) | 每轮账本基线、可 diff 的 keep/discard 实据 |
| `templates/outline.md` | 大纲 markdown 语法示例(入口 C 直接改它) | 写页纲/起稿时 |
| `templates/slides.json` | 可直接修改的 JSON 骨架 | JSON 精控起稿时 |
| `CHANGELOG.md` | v1.0→v2.6 全部版本历史与吸收来源 | 想知道某能力从哪来时 |

---

## 8. 明确不做(糟粕边界,防特性膨胀)

评估新特性时先对照此表;新增"不做"比新增"做"更需要理由。

| 不做 | 原因 |
|---|---|
| DOM→OOXML 的 HTML-first 渲染路线 | 笔格为此单列 30KB 分片,本地 skill 收益配不上工程量;结构化直出 + 图片版已覆盖两端 |
| 逐对象动画 | python-pptx 需裸 OOXML,易腐败难维护;交付原则"动画可丢不损内容" |
| 实时协同 / 云服务 / 积分计费 | 服务端产品形态,本地场景无对应物 |
| 浏览器端组装 OOXML | 非浏览器运行环境;python-pptx / pptxgenjs 已覆盖 |
| mock / 假图 / 假数据兜底 | 违反铁律,永不做 |
| 复杂公式的文本近似 | 用 tex2png 出图;mathtext 解析失败的式子宁可拆页,不上粗糙近似 |
| 实时图片搜索 API(需要 key 的图库) | 本地 skill 无 key 环境;维基共享资源(无需 key)已覆盖;商业图库留作可选扩展 |
