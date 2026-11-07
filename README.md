# wut-ppt — 武汉理工大学行政汇报 PPT 生成技能

以**武汉理工大学PPT模板**为固定骨架，内置**高校行政公文写作规范知识库**与**用户习惯自主演进引擎**，由 **AI 逐页设计正文 SVG**，经 PPT Master 转换为原生 DrawingML 后嵌回模板。面向党建、学工、教务、科研、行政、后勤等职能部门的工作汇报、述职述廉、年度总结、项目申报与评估答辩。

## 核心原则

```text
封面底版固定，正文版式不重；
公文文体严谨，对仗结论先行；
正文全面铺满，箭头必须表意；
自主学习反馈，技能持续进化。
```

- **文体定位**：汇报文体，不是阅读文体；让领导在 10~15 分钟内抓住要点、记住亮点、认同结论。
- **五条铁律**：工整优先于文采（句式同构/字数对齐）、稳妥优先于创新（体制内词库）、结论先行（拒中性标题）、虚实结合（排比金句+实干数据）、闭环结构（三点式成效/问题/计划严格对应）。
- **固定骨架**：封面、目录、章节页、红色页标题、底部蓝条和校徽必须来自学校定制模板。
- **正文多样**：SVG 只绘制安全区 `x=40–1240，y=76–688`，正文 Layout ID、几何签名、宏观构图全稿唯一。
- **彻底告别大片空白**：正文横向覆盖率≥90%、纵向≥85%，12×6 网格最大连续空白≤6格，空白必须用真实子要点、指标和任务消除。
- **箭头有因有果**：必须具备真实业务交接/映射语义，表格泳道无业务流转严禁乱加箭头。
- **自主演进闭环**：通过用户口头意见录入或对比用户手工改稿前后 PPTX，自动提炼习惯词汇与风格画像，越用越懂用户。

---

## 工作流

```text
源材料
→ 载入用户画像与公文写作规范
→ 全稿正文视觉地图规划（八字对仗/三点式/唯一构图）
→ AI 逐页手写正文 SVG（全幅面板/行政图风）
→ 公文合规审查（避讳词/中性标题）+ SVG 兼容性 + 版式与箭头审查
→ 原生 DrawingML 转换与固定模板组装（固化主题色）
→ PowerPoint PNG 像素级空白检查
→ 全稿 contact sheet 人工终审
→ 用户反馈 / 改稿自主学习提炼（技能持续进化）
```

---

## 环境要求

- Python 3.10+
- `python-pptx`、`Pillow`
- PPT Master v6.7.3（执行 `python scripts/engine_setup.py` 自动下载到 `.engine/`，不污染 Git）
- 预览与导出：Windows + 本机 Microsoft PowerPoint

```bash
pip install python-pptx Pillow
python scripts/engine_setup.py
```

---

## 快速上手与命令速查

### 1. 初始化并注入用户画像

```bash
python scripts/wut_svg.py init 我的工作汇报 --dir projects
```

命令创建项目、提取模板资产，并在 `design_spec.md` 中自动注入当前用户的历史偏好档案。

### 2. 编写规划（大纲、视觉地图、模板配置）

编辑：
- `design_spec.md`：标题采用八字对仗金句（如“聚焦立德树人，夯实育人根基”），段首三点式起头，每页规划唯一构图与几何签名。
- `spec_lock.md`：锁定颜色、字号和 dense 节奏。
- `deck.json`：固定封面标题、汇报人、章节名与正文标题列表。

参考权威知识库：
- 核心文体与公文武器词库：[`references/content_writing_guide.md`](references/content_writing_guide.md)
- 开放式构图选型：[`references/composition_library.md`](references/composition_library.md)
- 专有名词词典：[`references/terminology.md`](references/terminology.md)

### 3. AI 逐页手写正文 SVG

正文写入 `svg_output/NN_页面名.svg`，只绘制正文区，不包含封面底版。文本单元≥10、图形元素≥18，字号显式设置且≥16px（12pt）。

原生预设形状片段调用：
```bash
python scripts/wut_svg.py preset render downArrow --id flow-1 --frame 100 300 120 90 --fill '#DCE6F5' --stroke none
```

### 4. 技术、公文规范与版式检查

```bash
python scripts/wut_svg.py check projects/我的工作汇报_目录
```

检查包含：
- 公文合规：绝对避讳词（“失败/错误/严重问题/搞砸/经费不足/人手不够/咱们”等）拦截、中性标题拦截。
- 版式唯一性：Layout ID、几何签名查重、表格无业务语义箭头拦截。
- 技术与字号：12pt 下限、SVG 语法规范、文本图形元素基线。

### 5. 导出并嵌回固定模板

```bash
python scripts/wut_svg.py export projects/我的工作汇报_目录
```

自动完成原生 DrawingML 转换，克隆张林模板封面、目录、章节与母版，注入正文并固化 RGB 主题色，防止串色。

### 6. 一键视觉验证

```bash
python scripts/wut_svg.py verify projects/我的工作汇报_目录
```

自动导出逐页高精度 PNG，生成 `contact_sheet.png` 全稿总览，并运行像素级空白分析（覆盖率不达标直接报错）。

### 7. 用户自主学习与技能演进

技能具备两种自我学习模式：

- **日常要求录入**：
  ```bash
  python scripts/wut_svg.py feedback "校领导汇报更偏好八字对仗，成效多用同比增幅，把数据平台统一称为全域数智中台"
  ```
- **对比改稿自主学习**（将用户自己手动在 PowerPoint 修改后的成品喂给技能反思）：
  ```bash
  python scripts/wut_svg.py learn projects/我的汇报/exports/原版.pptx 用户修改版.pptx
  ```
  技能自动比对文本替换、词汇升格与删改习惯，沉淀至 `user_profile.json` 与 `references/user_learned_preferences.md`，下次生成自动生效！
- **查看当前用户专属规则**：
  ```bash
  python scripts/wut_svg.py profile
  ```

---

## 项目结构与知识库

```text
wut-ppt/
├── SKILL.md                          # 技能入口（核心原则、七步工作流、验收清单）
├── README.md                         # 本文档
├── assets/template.pptx              # 学校定制模板（固定封面/目录/章节/母版）
├── user_profile.json                 # 当前用户习惯与规则沉淀画像
├── scripts/
│   ├── engine_setup.py               # PPT Master v6.7.3 下载安装
│   ├── wut_svg.py                    # 统一 CLI 工具
│   ├── assemble_deck.py              # 固定模板克隆与原生正文注入
│   ├── check_content_compliance.py   # 高校公文合规、避讳词与中性标题拦截
│   ├── deck_visual_audit.py          # 版式唯一性与箭头业务语义检查
│   ├── check_visual_density.py       # SVG 元素与 PNG 像素级连续空白硬检查
│   ├── learn_from_feedback.py        # 用户习惯录入与 PPTX 改稿差异自主学习引擎
│   ├── check_deck.py                 # PPTX 终稿自检
│   ├── export_preview.py             # 逐页高清 PNG 导出
│   └── make_contact_sheet.py         # 全稿缩略图总览生成
└── references/
    ├── content_writing_guide.md      # 高校行政公文五条铁律与核心武器词库
    ├── user_learned_preferences.md   # 自主学习引擎提炼的用户专属规则库（动态进化）
    ├── design_rules.md               # 全幅行政图风与填满正文规则
    ├── outline_contract.md           # 正文规划、deck.json 与 SVG 编写契约
    ├── composition_library.md        # 正文开放式构图候选与选型指南
    └── terminology.md                # 学校专名与术语词典
```

---

## 信息安全保证

- `_work/`、`.engine/`、`projects/`、预览图均被 `.gitignore` 严格忽略，绝不上传到公开仓库。
- 测试与示例一律使用脱敏虚拟表述。
