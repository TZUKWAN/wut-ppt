---
name: wut-ppt
description: 武汉理工大学行政材料与高信息密度 PPT 生成技能。张林定制模板的封面、目录、章节页、红色页标题和底部母版固定不变；AI 只用完整 SVG 设计正文内容区，经 PPT Master 转原生 DrawingML 后嵌回模板。正文宏观版式不得重复、正文必须占满不得留下连续空白、箭头必须有业务语义、全文不小于12pt，导出后必须通过像素级空白检查与全稿复检。触发词：做汇报PPT、工作汇报、进展汇报、阶段总结、行动计划汇报、武汉理工、智链、信息可视化。
---

# 武汉理工大学行政汇报 PPT 技能（固定模板＋SVG 正文）

**核心路线：源材料 → 全稿正文视觉地图 → AI 手写正文 SVG → 原生 DrawingML → 嵌回张林固定模板 → 像素级密度检查与全稿复检。**

## 不可违反的质量规则

1. **固定骨架不可变**：封面、目录、章节页、正文红色页标题、底部蓝条和校徽必须克隆 `assets/template.pptx`，不得用 SVG 重新设计。
2. **多样化只发生在正文区**：同一套 PPT 的 `Layout ID`、`Geometry signature`、`Macro composition` 均不得重复；换内容、颜色或图标不算新版式。
3. **正文必须占满**：安全区 `x=40–1240，y=76–688` 必须充分铺开，禁止连续空白区域，禁止以“呼吸感”“大留白”为设计目标。
4. **空白靠真实内容消除**：用子要点、任务、指标、交付物、输入输出和关系节点补足，不得只拉高空容器或添加无意义装饰。
5. **箭头必须有业务语义**：每个箭头都必须能在 `Arrow semantics` 中说明“来源→目标＋交接/映射/转化/支撑关系”；表格、泳道没有真实交接时不得加箭头。
6. **每页一个主视觉命题**：先确定叙事任务和内容关系，再决定构图；禁止默认卡片网格和简单中心圆加四周漂浮文字。
7. **行政图风必须饱满**：使用有信息职责的实色面板、粗引导线、金色按钮、项目符号和完整解释；避免只有细线、淡色节点和大片白底的示意图。
8. **全文不小于 12pt**：1280×720 画布中任何显式 `font-size` 不得小于 16px；正文建议 18–22px。
9. **信息密度有下限**：每个正文页至少 10 个文本信息单元和 18 个可见图形元素，这是下限而非凑数目标。
10. **全稿目检不可省**：固定结构页可以重复，正文页不得近义同构，也不得出现局部空白、串色、遮挡和断词。

## 工作流

### ① 准备引擎

```bash
python scripts/engine_setup.py
```

### ② 初始化并规划

```bash
python scripts/wut_svg.py init 项目名 --dir 项目父目录
```

完成：

- `design_spec.md`：正文内容、关系、唯一构图和几何签名。
- `spec_lock.md`：颜色、字体和结构锁，正文页面节奏只用 dense。
- `deck.json`：固定封面、汇报人、章节和正文页标题，页序与 SVG 一致。

每个正文页必须填写 `Layout ID`、`Geometry signature`、`Macro composition`、`Reading direction`、`Visual center`、`Rendering language`、`Distinction`。**使用箭头时必须另填 `Arrow semantics`。**

构图从 `references/composition_library.md` 选择或定制，但任何构图都必须展开到整个正文区，不能留出整片空白。

### ③ AI 逐页手写正文 SVG

- `svg_output/NN_页面名.svg` 只包含正文，不包含封面、目录、章节页、页标题、底部蓝条、校徽或整页品牌背景。
- 主 Agent 逐页连续创作，不用 Python 批量生成页面，不委托子 Agent 分页设计。
- 动笔前回答：本页与此前正文有何本质区别？内容如何铺满整个正文区？每个箭头表达什么真实关系？
- 中文用 `<text>` + `<tspan>` 主动换行；不得靠缩小字号处理溢出。
- 每个逻辑模块使用唯一语义 `<g id>` 和 `data-pptx-bounds`。
- 原生预设通过 `wut_svg.py preset` 生成；预设箭头同样受语义审查。

### ④ 技术检查与固定模板组装

```bash
python scripts/wut_svg.py check 项目目录
python scripts/wut_svg.py export 项目目录
```

`check` 会拦截：结构错误、SVG 不兼容、文字小于 12pt、版式重复、箭头缺少语义、文本或图形信息单元不足。

`export` 继续完成原生 DrawingML 转换、张林模板克隆、正文嵌入、主题色固化和 PPTX 自检。

### ⑤ 一键视觉验证

```bash
python scripts/wut_svg.py verify 项目目录
```

该命令自动导出逐页 PNG、生成 `contact_sheet.png`，并执行像素级正文检查：

- 正文横向铺展不得低于 90%，纵向不得低于 85%。
- 正文划分为 12×6 网格，最大连续空白区域不得超过 6 格。
- 自动跳过固定封面、目录和章节页，只检查正文。

任何一项失败都必须回到内容规划或 SVG 修改，不得用“行政 PPT 也需要留白”解释通过。

### ⑥ 人工终审

打开 `contact_sheet.png` 和关键单页：

- 固定模板是否完全一致、无串色。
- 正文是否真正饱满，是否仍有局部整片空白。
- 是否有无意义箭头、简单中心圆、漂浮标签或换皮卡牌。
- 正文是否近义同构，是否存在遮挡、断词和小于 12pt 的文字。

## 文件职责

```text
assets/template.pptx              张林定制模板（固定骨架唯一来源）
scripts/engine_setup.py           下载固定版本 PPT Master 引擎
scripts/wut_svg.py                init / preset / check / export / verify 统一入口
scripts/assemble_deck.py          克隆固定模板并嵌入原生正文、固化主题色
scripts/deck_visual_audit.py      版式唯一性与箭头语义审查
scripts/check_visual_density.py   SVG 信息单元与 PNG 空白区域硬检查
scripts/check_deck.py             PPTX 占位符、字号、色值、空页检查
scripts/export_preview.py         PowerPoint COM 逐页 PNG 导出
scripts/make_contact_sheet.py     全稿缩略图总览
references/design_rules.md        正文行政图风与全页填满规则
references/outline_contract.md    正文规划及 SVG 编写契约
references/composition_library.md 正文开放式构图候选与选型规则
references/terminology.md         武理工术语和专名约束
```
