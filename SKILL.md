---
name: wut-ppt
description: 武汉理工大学行政材料与高信息密度 PPT 生成与自主进化技能。张林定制模板的封面、目录、章节页、红色页标题和底部母版固定不变；严格遵循高校公文五条铁律（工整优先、稳妥优先、结论先行、虚实结合、闭环结构）；AI 逐页设计全幅正文 SVG，经 PPT Master 转原生 DrawingML 后嵌回模板；具备自动化公文避讳词拦截、版式唯一性与像素级空白检查；支持从用户日常反馈及 PPTX 改稿记录中自主提炼偏好并持续迭代进化。触发词：做汇报PPT、工作汇报、进展汇报、阶段总结、行动计划汇报、武汉理工、智链、信息可视化、行政PPT。
---

# 武汉理工大学行政汇报 PPT 技能（固定模板＋SVG 正文＋公文文体＋自主进化）

**核心定位：高校汇报文体，不是阅读文体**——让校领导/评委在 10~15 分钟内抓住要点、记住亮点、认同结论。
**完整闭环：源材料 → 公文规范提炼 → 用户画像注入 → AI 手写正文 SVG → 原生 DrawingML → 嵌回固定模板 → 像素级密度与公文审查 → 用户反馈/改稿自主学习迭代。**

---

## 一、高校公文内容写作五条铁律（不可违反）

1. **工整优先于文采**：同层级标题句式同构、字数绝对对齐、节奏一致，优先采用四字/六字/八字对仗（如“聚焦立德树人，夯实育人根基”），比抒情辞藻重要得多。
2. **稳妥优先于创新**：用词严格在体制内公文武器库范围内（参考 `references/content_writing_guide.md`），不生造、不猎奇、不口语化。动作词主动升格（做→推进、搞活动→搭建载体、查问题→对标查摆、要钱→经费保障）。
3. **结论先行**：每页标题即政绩结论与核心论断，严禁“XX工作概况/XX建设总结”式的中性无观点标题。
4. **虚实结合**：小标题是虚的对仗排比金句，下方正文/图表是实的数据注脚、落地任务、具体举措与责任机制。
5. **闭环结构**：高校行政偏好**三点式/四点式**；成效、问题、计划严格 1:1 对应，首尾呼应，听众能清晰复述出“三个点”。

---

## 二、视觉排版不可违反的质量底线

1. **固定骨架不可变**：封面、目录、章节页、正文红色页标题、底部蓝条和校徽必须克隆 `assets/template.pptx`，严禁用 SVG 篡改或替换。
2. **多样化只发生在正文区**：同一套 PPT 正文的 `Layout ID`、`Geometry signature`、`Macro composition` 均不得重复；只改文字、颜色或局部形状仍算重复。
3. **正文必须占满，痛恨大片空白**：正文安全区 `x=40–1240，y=76–688` 必须充分铺开，禁止连续空白区域，禁止以“留白”“呼吸感”为由留出空洞。
4. **空白靠真实业务内容消除**：用子要点、任务、指标、交付物、输入输出和关系节点补足，严禁仅机械拉大容器或加无意义色块。
5. **箭头必须有业务语义**：每个箭头都必须能在 `Arrow semantics` 说明“来源→目标＋交接/映射/转化/支撑关系”；表格、泳道无真实交接时严禁加箭头。
6. **行政图风必须饱满稳重**：实色面板、粗引导线、金色按钮、醒目项目符号与完整解释；坚决摈弃细线、浅淡节点加大片白底的简陋示意图。
7. **全文不小于 12pt**：1280×720 画布中任何显式 `font-size` 不小于 16px（12pt）；正文建议 18–22px。
8. **密度双重硬门槛**：SVG 文本单元≥10、图形元素≥18；PNG 渲染后横向铺展≥90%、纵向≥85%，12×6 网格最大连续空白≤6 格。
9. **绝对避讳红线**：严禁口语（搞定/做/咱们）、负面定性（失败/错误/严重问题）、绝对示弱（人手不够/缺钱/没人干）。问题必须写成“发展中的问题、前进中的不足、对照更高标准的差距”。

---

## 三、工作流（七步闭环）

### ① 准备引擎（每个技能副本一次）
```bash
python scripts/engine_setup.py
```

### ② 拆材料、加载用户画像并初始化
保留源材料真实数据和专名，初始化项目：
```bash
python scripts/wut_svg.py init 项目名 --dir 项目父目录
```
- 自动提取模板媒体，并在 `design_spec.md` 头部自动加载并注入当前用户的历史偏好（来自 `user_profile.json`）。

### ③ 规划公文大纲与正文视觉地图
依据 `references/content_writing_guide.md`、`references/outline_contract.md` 和 `references/composition_library.md`：
- **标题**：八字对仗或动宾同构，结论先行。
- **正文**：三点式起头（一是/二是/三是），数据突出增量百分比，层级按国→省→校排序。
- **问题与计划**：问题控制在 2~3 条客观差距，与后文计划 1:1 闭环对应。
- **版式地图**：每页填报唯一 `Layout ID`、`Geometry signature`、`Macro composition`，有箭头必填 `Arrow semantics`。

### ④ AI 逐页手写正文 SVG
- `svg_output/NN_页面名.svg` 只绘制正文安全区（`x=40–1240，y=76–688`）。
- 动笔前自检：本页构图与此前正文是否有本质区别？内容是否铺满？箭头是否承载真实业务流？
- 中文用 `<text>` + `<tspan>` 主动排版换行，严禁靠缩小字号塞字。
- 原生预设形状调用：`python scripts/wut_svg.py preset render downArrow ...`。

### ⑤ 技术与公文合规检查、原生导出与模板组装
```bash
python scripts/wut_svg.py check 项目目录
python scripts/wut_svg.py export 项目目录
```
- `check` 执行五重关卡：项目结构 → SVG 兼容性 → 12pt 字号 → 公文合规与避讳词扫描 → 版式唯一性与箭头语义 → 静态信息密度。
- `export` 执行 DrawingML 原生转换，并调用 `assemble_deck.py` 克隆固定模板、注入原生正文、固化 RGB 色值。

### ⑥ 一键视觉验证与人工终审
```bash
python scripts/wut_svg.py verify 项目目录
```
- 自动导出 PNG、生成 `contact_sheet.png`，并对正文进行像素级铺展与空白网格检查（覆盖率未达标直接报错拦截）。
- 打开总览图人工复查：模板与颜色是否统一无串色、正文有无近义同构、有无局部空白或无意义箭头。

### ⑦ 用户反馈吸收与技能自主演进（Autonomous Learning Loop）
这是让技能“越用越懂当前用户”的核心闭环：
- **日常偏好即时学习**：当用户提出特定偏好、领导习惯或意见时，直接录入自学习引擎：
  ```bash
  python scripts/wut_svg.py feedback "张校长汇报必须突出立德树人，成效数据多用百分比，不要用英文缩写"
  ```
  引擎自动提炼特征并沉淀至 `user_profile.json` 与 `references/user_learned_preferences.md`。
- **用户改稿差异自主反思**：当用户在 PowerPoint 中手动调整了导出的 PPTX 后，对比学习：
  ```bash
  python scripts/wut_svg.py learn 项目目录/exports/原版.pptx 用户修改后的文件.pptx
  ```
  引擎自动比对文本替换、词汇升格与删改习惯，将其转化为下一次生成时的优先规则！
- **查看当前用户画像**：
  ```bash
  python scripts/wut_svg.py profile
  ```

---

## 四、文件清单与知识库职责

```text
assets/template.pptx                  张林定制模板（固定封面/目录/章节/母版唯一来源）
user_profile.json                     当前用户专属习惯、词汇替换与风格偏好画像
scripts/
├── engine_setup.py                   固定版本 PPT Master v6.7.3 依赖安装
├── wut_svg.py                        init / preset / check / export / verify / feedback / learn / profile 统一入口
├── assemble_deck.py                  克隆固定模板、注入原生正文、固化 RGB 色值
├── check_content_compliance.py       高校公文合规、避讳词拦截、中性标题审查
├── deck_visual_audit.py              正文版式唯一性、SVG 几何查重与箭头语义审查
├── check_visual_density.py           SVG 元素密度与 PNG 像素级连续空白硬检查
├── learn_from_feedback.py            用户意见录入、改稿比对与自主演进引擎
├── check_deck.py                     PPTX 占位符、字号下限、色值合规检查
├── export_preview.py                 PowerPoint COM 逐页导出 PNG
└── make_contact_sheet.py             生成全稿 contact_sheet 缩略图总览
references/
├── content_writing_guide.md          高校行政公文写作五条铁律与核心武器词库
├── user_learned_preferences.md       自学习引擎沉淀的用户偏好规则库（动态演进）
├── design_rules.md                   全幅行政图风与填满正文规则
├── outline_contract.md               正文规划、deck.json 与 SVG 编写契约
├── composition_library.md            正文开放式构图候选与选型指南
└── terminology.md                    武汉理工大学专名与术语词典
```
