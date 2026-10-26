# CHANGELOG · ppt-master-pro

## v2.6.1 · 落地校验与修复(2026-10-26)

v2.6 的新增五类页面此前从未做过视觉复核,本轮逐页渲染(LibreOffice → PDF → PNG)后定位并修掉 8 个真实缺陷,补上缺失的语法文档。

**修复(全部来自渲染实证,不是猜)**:
- `quadrant` 象限图**几何错误**:右侧两格起点用了区域右边界而非中点,整片画到画布外;竖轴起止点同为 y=中点导致长度为零、完全画不出来。改为以区域正中为原点,横竖轴贯穿全区域形成十字,四条轴标签贴各自轴的端头
- `matrix` 对比矩阵**:左上角空格露出 python-pptx 默认表格样式的品牌蓝**,改为显式填主色,与表头行连成一条
- `architecture` 架构图**箭头错位**:原按下标算固定偏移(`mid + (k-1.5)*1.4`),单组件层会多画一根插在框边上的游离箭头。改为按各框中心落点,组件多于 3 个时取首/中/末
- `dashboard` 仪表盘**多余白线**:末行之后仍在画分隔线(条件 `r < len(stats)//per_row` 在正好一行时恒真);行高改为设上限并把剩余空间上下平分,不再顶在页顶
- `cards` 卡片**撑成三个空盒子**:卡片高度改为按文案长度自适(1.9–3.4in)并垂直居中
- `quote` 引号装饰**字形崩坏**:微软雅黑 96pt 粗体的 “ 渲染成两道斜杠;改走 Georgia 110pt,并把落款 y 改为跟随引文行数自适应(原来硬编码 5.3in 中间空一大块)
- `generate_pptx.py --doctor`**永远跑不起来**:`from pptx import ...` 在模块顶部,缺依赖时导入即崩,而 `--doctor` 的职责正是诊断缺依赖。改为软导入 + 数值垫片,缺依赖时自检照常输出并给出"装给哪个解释器"的精确命令
- **坏页跳过会留下白页**:构建器抛异常时只 `continue`,空页仍在成品里。改为 `_drop_last_slide()` 摘掉(操作 sldIdLst + drop_rel),页码与页脚按"真正建成的页"重新编号
- `qa_report.py`**溢出判定全是误报**:估算式 `chars*fs*0.9/(width_in*72)*1.2` 量纲混乱,17 页每页都报"估 1.4in > 框 0.3in";空矩形(自动图形本无文字)被当成"空文本框"。改为与 `generate_pptx` 同一套字宽推算法,空文本框只认 `shape_type==17`,单字符装饰不参与估算
- `qa_report.py`/`extract_theme.py`/`pptx2outline.py`/`build_image_deck.py` 缺依赖时抛裸 traceback(前者连 `import sys` 都没导入),统一改为一行可执行提示

**新增**:
- `templates/slides_all_types.json` — 全 17 种页面类型的可运行样例(每页带 notes,演示"备注不是正文复读"),原来只有 8 类有样例,另 9 类无参考
- SKILL.md 补 §「v2.6 新增 5 类页面语法」字段表(必填/可选/语义要点)——此前 `matrix/quadrant/architecture/cards/dashboard` 全文只被提过一句名字,字段名无任何文档,调用方只能靠读源码猜

**文档一致性**:
- 版本号统一 v2.5 → v2.6.1(SKILL.md frontmatter、README 目录树、文件表)
- 数量口径统一 12 类型 × 7 主题 → 17 类型 × 12 主题(4 处残留)
- 主题清单补全 12 套 + 说明 `style:"auto"` 走 tags 关键词重合度匹配
- 合并重复的「整合来源」段落(两行内容重复,保留并合二为一)

**回归结果**(Python 3.14 + python-pptx 1.0.2 + LibreOffice 25):
- 17 类型全量 deck:0 未知类型 / 0 越界 / 0 溢出警告,17 页逐页渲染复核通过
- 带备注模板 deck:qa_report 综合 100.0(content/design/coherence 全 100),PASS
- md 直译、json2html、pptx2outline、extract_theme、--doctor 全部跑通

## v2.6 · guizang + codex-ppt + 全量扩展(2026-10-24)

**吸收自**:op7418/guizang-ppt-skill(内容风格双轨:Style A 电子杂志×电子墨水 / Style B 瑞士国际主义;内置排练与演讲者模式;多平台封面)、kahming/codex-ppt(Markdown-first 驱动全流程、四阶段生成循环、角色分离 concept/design/image/shell、三层双轨 style/layout/themes)。

**新页面类型(12→17)**:
| 类型 | 用途 | 原生实现 |
|---|---|---|
| matrix | 对比矩阵(表头+行标签+焦点格高亮) | python-pptx 表格 + 单元格填充 |
| quadrant | 象限图(坐标轴交叉+点落位+焦点标注) | 形状+连接线+文本标注 |
| architecture | 架构图(分层列表+层内横排+层间箭头) | 形状+连接线+箭头 |
| cards | 组合卡片行(heading+desc+色块图标区,浅底+阴影) | 形状+阴影 |
| dashboard | 大数字仪表盘(值40-52pt+单位+标签,行间分隔线) | 文本+连接线 |

**新主题(7→12)**:finance_green(金融墨绿)、medical_blue(医疗青)、law_burgundy(法律酒红)、education_teal(教育培训青)、luxury_dark(奢华黑金)——每套带行业 tags 参与 auto 匹配。

**新脚本**:
- `fetch_image.py` — 维基共享资源配图抓取(无需 API key),替代手动找图
- `pptx2outline.py` — 现有 PPTX → outline.md 结构反解析(讯飞 parse-pptx 本地化),实现"拿到别人 deck → 改风格/内容重出"
- `qa_report.py` — pptx → 机器可读 QA 报告 JSON(PPTEval 三维程序化部分),实验账本的确定性基线

**文档升级**:
- 内容双轨(叙事轨 story / 事实轨 fact)——Phase 0 定轨
- 页面脚本(codex-ppt):页纲的页级细化——主张句+版式+视觉焦点+图片需求
- 角色分离(concept/design/image/shell):六阶段天然有,明确写出
- 三层双轨(风格轨/版式轨/主题轨):模板双轨升级
- 排练模式:网页 deck 自带备注+计时器(N 呼出/R 归零),交付附排练指引
- §8 明确不做新增:需要 key 的图库(维基已覆盖,商业图库留作扩展)

## v2.5 · 精华吸收 + 去糟粕整理(2026-10-23)

**吸收自**:zarazhangrui/frontend-slides(视觉风格发现:3 个单页小样代替语言描述审美;零依赖单文件 HTML)、Slidev(演讲者视图:备注+计时器;KaTeX 公式)、pptx-from-layouts(母版布局填充——我们已有,验证方向)。

- **新增 `scripts/tex2png.py`**:LaTeX → 透明 PNG(matplotlib mathtext,无需 LaTeX),学术模式配套;公式多的 deck 逐式出图后用 image 页承载
- **演讲者计时器**:网页 deck 按 N 呼出备注时自动计时,按 R 重置(Slidev presenter view 精简版)
- **风格小样选型**:用户说不清审美时,用现成管线出 3 个单页小样(安全/大胆/出格各一)让其挑,代替语言描述
- **HTML deck → PDF 配方**:Playwright 逐页截图 → build_image_deck --pdf(复用现有工具,不新增依赖)

**去糟粕(明确不做,防特性膨胀)**:
- DOM→OOXML 的 HTML-first 渲染路线(笔格为此单列 30KB 分片,本地 skill 收益配不上工程量)
- 逐对象动画(python-pptx 需裸 OOXML,易腐败难维护;交付原则"动画可丢不损内容")
- 实时协同/云服务/积分计费(服务端产品形态,本地场景无对应物)
- 浏览器端写 OOXML(我们不是浏览器环境;python-pptx/pptxgenjs 已覆盖)
- mock/假图/假数据兜底(违反铁律,永不做)

## v2.4 · Gamma 全流程取证(2026-10-23)
风格档案持久化 `ppt-profile.json`(环境变量→spec 同目录→~/.ppt-master-pro/,spec 显式永远优先;`extract_theme.py --save-profile` 提取即存档)· 质量三档 lite/standard/max 决定质检深度 · 模板双轨显式化(换肤 themeId vs 保结构 templateId)· 页数=软目标(偏离须记账)· 版本快照 `<stem>_rN.pptx` · 光栅化兜底 · 多模型分工。

## v2.3 · autoresearch × D-PPT-skill(2026-10-23)
第三输出模式·图片版 deck(`build_image_deck.py`:整页 PNG → PPTX/PDF 封装 + 交付三查,`--expect` 页数校验)· 不假图兜底(生图失败明确停止)· QA 实验账本(每轮 `R<N>: 问题→修复→复验→keep/discard`,≤2 轮)· 断点续跑(分批/生图逐页落盘)。

## v2.2 · 商业产品逆向(WPS/讯飞/笔格/歌者/轻竹,2026-10-23)
一键美化 `extract_theme.py`(参照 PPTX → 三角色主题,面积加权+字色并票,闭环 4/4 角色命中)· `"style":"auto"` 行业词打分选主题(0.6×命中+0.4 基础)· 长文档分块大纲(>30 页分块摘要→合并→幽灵 deck 定去留)· 讲稿双产出 `--script` · 单页重绘说明。

## v2.1 · academic-pptx-skill × ultimate-ppt-master(2026-10-23)
双输出模式(同一大纲 → PPTX + 杂志风网页 deck `json2html.py`)· 幽灵 deck 测试(只读主张句应讲完整个故事)· 学术模式(exhibit 纪律/in-text 引用+References 页/结论页收尾)· 设计规格锁定(design_spec.md + relock)。

## v2.0 · PPTAgent × Presenton × presentation-ai(2026-10-23)
三入口模式(从零/编辑式/大纲直译)· `md2json.py` 大纲 markdown 直译 · tone×verbosity×audience 三参数 · PPTEval 三维评估(Content/Design/Coherence)· 视觉审查-修复循环(≤2 轮只修 fail 页)· 页面功能类型学 · 检索接地 · `--doctor` 自检 · 主题 5→7 套。

## v1.0 · 初版(2026-10-23)
整合 hugohe3/ppt-master、gnuhpc/ppt-master-plus、Offwhite-Del 294 技法、editable-ppt-master、pptx-generator、presentations:pptx、anti-defensive-writing、ui-ux-pro-max:六阶段流水线 + JSON 生成器(12 类型×5 主题)+ 说人话体系 + 去 AI 味设计规范。
