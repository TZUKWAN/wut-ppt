# wut-ppt — 武汉理工大学行政汇报 PPT 生成技能

以**武汉理工大学PPT模板**为固定骨架，由 **AI 逐页设计正文 SVG**，经 PPT Master 转换为原生 DrawingML 后嵌回模板。面向武汉理工大学行政材料、工作汇报、进展汇报、行动计划和高密度信息可视化页面。

## 核心原则

```text
封面和底版固定，正文版式不重复；正文必须占满，箭头必须表意。
```

- 封面、目录、章节页、红色页标题、底部蓝条和校徽必须来自张林模板。
- SVG 只绘制正文安全区 `x=40–1240，y=76–688`。
- 正文 Layout ID、几何签名、宏观构图只能出现一次。
- 禁止连续空白，禁止用“呼吸感”解释空疏页面。
- 空白必须用真实子要点、解释、任务、指标、交付物和关系节点消除。
- 表格、泳道没有真实交接或映射时不得添加箭头。
- 所有显式文字 `font-size ≥ 16px`，即 **PPT ≥12pt**。
- 最终正文和固定模板均为原生可编辑对象。

## 工作流

```text
源材料
→ 全稿正文视觉地图（design_spec + spec_lock + deck.json）
→ 每页候选构图与跨页差异审查
→ AI 逐页手写正文 SVG
→ 版式、箭头语义和静态密度检查
→ 原生 DrawingML 转换与固定模板组装
→ PowerPoint PNG 像素级空白检查
→ 全稿 contact sheet 人工终审
```

## 环境要求

- Python 3.10+
- `python-pptx`、`Pillow`
- PPT Master v6.7.3（安装命令自动下载到 `.engine/`）
- 逐页 PNG 预览：Windows + 本机安装 Microsoft PowerPoint

```bash
pip install python-pptx Pillow
python scripts/engine_setup.py
```

## 快速开始

### 1. 初始化并规划

```bash
python scripts/wut_svg.py init 我的汇报 --dir projects
```

编辑：

- `design_spec.md`：正文内容、关系、唯一构图、几何签名；有箭头时另填 `Arrow semantics`。
- `spec_lock.md`：色彩、字体和结构锁，正文只用 dense。
- `deck.json`：封面标题、汇报人、章节和正文页标题，页序与 SVG 一致。

详见 [`references/outline_contract.md`](references/outline_contract.md)、[`references/design_rules.md`](references/design_rules.md) 和 [`references/composition_library.md`](references/composition_library.md)。

### 2. AI 逐页手写正文 SVG

把每页正文写入 `svg_output/NN_页面名.svg`，不得包含固定封面、页标题或母版。每页至少 10 个文本信息单元和 18 个可见图形元素，且必须真实、有效。

原生 PowerPoint 预设形状：

```bash
python scripts/wut_svg.py preset render downArrow \
  --id flow-1 --frame 100 300 120 90 \
  --fill '#DCE6F5' --stroke none
```

### 3. 检查、转换与固定模板组装

```bash
python scripts/wut_svg.py check projects/我的汇报_日期目录
python scripts/wut_svg.py export projects/我的汇报_日期目录
```

检查失败条件包括：版式重复、箭头缺少来源→目标及业务语义、文字小于 12pt、文本或图形信息单元不足、SVG 不兼容。

### 4. 一键视觉验证

```bash
python scripts/wut_svg.py verify projects/我的汇报_日期目录
```

自动导出 PNG、生成全稿总览，并检查：

- 正文横向铺展 ≥90%、纵向 ≥85%。
- 12×6 网格最大连续空白 ≤6 格。
- 只检查正文，固定封面、目录和章节页自动跳过。

最后打开 `contact_sheet.png` 和关键单页，人工排除近义同构、局部空白、无意义箭头、遮挡和断词。

## 项目结构

```text
wut-ppt/
├── SKILL.md
├── README.md
├── assets/template.pptx
├── scripts/
│   ├── engine_setup.py
│   ├── wut_svg.py
│   ├── assemble_deck.py
│   ├── deck_visual_audit.py
│   ├── check_visual_density.py
│   ├── check_deck.py
│   ├── export_preview.py
│   └── make_contact_sheet.py
└── references/
    ├── design_rules.md
    ├── outline_contract.md
    ├── composition_library.md
    └── terminology.md
```

## 信息安全

- `_work/`、`.engine/`、预览目录均不进入 Git。
- 不上传真实行政材料、含姓名或内部数据的示例大纲、成品 PPT 和预览图。
- 验证与测试使用脱敏虚拟内容。

## 底层引擎与许可

SVG 检查、预设形状和 DrawingML 转换复用 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)（MIT License）。本技能负责武汉理工大学专属的内容规划、固定模板组装、正文版式多样性、行政信息密度和验收工作流。
