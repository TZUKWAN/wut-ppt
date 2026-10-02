# wut-ppt — 武汉理工大学行政汇报 PPT 生成技能

以**武汉理工大学PPT模板**为唯一载体（校徽蓝 `#00469A` + 金黄 `#FBC540` + 红 `#C00000`），用一份**大纲 JSON** 自动生成高密度、逻辑图风的行政汇报 PPT。

本仓库同时是一个 **Kimi Code / Claude Code 风格的 Agent 技能包**（含 `SKILL.md`），可直接被 AI 编程助手加载调用。

## 五条质量硬规则

1. **字号下限 12pt** —— 任何文字不得小于 12pt（脚本强制检查）
2. **每个要点独立形状包裹** —— 圆角卡 + 序号圆徽，禁止关键词漂浮
3. **视觉逻辑引导** —— 递进用箭头链、时间用时间轴、并列用连接线、映射用箭头
4. **高信息密度** —— 要点 = 引导语 + 完整解释句，每页 300–600 字
5. **零大片空白** —— 垂直流布局引擎自动把剩余空间分配给内容区，填满整页

## 快速开始

环境要求：Python ≥ 3.9；依赖 `python-pptx`；预览导出需要 Windows + 本机安装 PowerPoint（可选步骤）。

```bash
pip install python-pptx

# ① 按下方格式写一份大纲 JSON（references/outline_contract.md 有完整契约和页面配方）
# ② 生成 PPT
python scripts/build_deck.py 大纲.json 输出.pptx

# ③ 机器自检（占位符残留 / 字号下限 / 色值合规 / 空页）
python scripts/check_deck.py 输出.pptx

# ④ 导出每页 PNG，逐页视觉自检（需本机 PowerPoint）
python scripts/export_preview.py 输出.pptx _preview
```

## 大纲 JSON 格式

```json
{
  "title": "封面主标题",
  "speaker": "汇报人",
  "chapters": [
    {"name": "章节名",
     "pages": [
       {"title": "页面标题",
        "zones": [
          {"type": "lead", "text": "导语段……"},
          {"type": "cards", "cards": [
            {"head": "维度一", "items": [
              {"lead": "引导语", "text": "完整解释句……"}
            ]}
          ]},
          {"type": "note", "text": "结论条……"}
        ]}
     ]}
  ],
  "closing": {"line1": "以上汇报，敬请批评指正", "line2": "汇报人：XXX"}
}
```

- 封面 / 目录 / 章节页 / 封底**自动生成**，只需写正文页的 `zones`
- 文字标记：`**蓝粗强调**`、`[[红粗关键数字]]`

### 12 种 zone 版式

| zone | 用途 | zone | 用途 |
|---|---|---|---|
| `lead` | 页首导语段 | `badge_grid` | 徽章网格（N×M 亮点清单） |
| `cards` | 列卡片（多维并列，最常用） | `shape_row` | 形状横排（组织/角色 + 连接线） |
| `flow_chain` | CHEVRON 箭头链（递进逻辑） | `chips` | 六边形标签组 + 说明 |
| `timeline` | 时间轴（分阶段安排） | `kv_rows` | 键值行（逐条定义） |
| `compare_rows` | 对象行（维度→措施映射） | `banner` | 页首主题横幅 |
| `table` | 数据表格 | `note` | 页脚结论条 |

字段细节与 7 种**页面配方**（现状判断页 / 问题诊断页 / 组织措施页 / 制度激励页 / 时间安排页 / 计划总览页 / 阶段实施页）见 [`references/outline_contract.md`](references/outline_contract.md)。

## 项目结构

```
wut-ppt/
├── SKILL.md                  # Agent 技能入口（触发词/硬规则/工作流/验收清单）
├── assets/
│   └── template.pptx         # 张林定制模板（唯一载体，5 样板页 + 母版）
├── scripts/
│   ├── build_deck.py         # 生成器：zone 垂直流布局引擎
│   ├── check_deck.py         # 机器自检
│   └── export_preview.py     # PNG 导出（PowerPoint COM）
└── references/
    ├── outline_contract.md   # 大纲 JSON 契约 + 页面配方
    ├── design_rules.md       # 设计系统（色板/字号/形状库）+ 失败教训分析
    └── terminology.md        # 武汉理工术语词典
```

## 设计规范摘要

- **颜色**：主蓝 `#00469A` / 金黄 `#FBC540` / 强调红 `#C00000` + 黑白灰中性色，禁止其他颜色
- **字体**：全文微软雅黑；正文 12pt 起，列头 12.5–13pt，横幅 14pt
- **形状**：圆角矩形（卡片）、圆形（徽章/节点）、CHEVRON（递进链）、六边形（标签）、平行四边形（对象签）、金色箭头（映射）——禁止大色块椭圆/菱形
- **模板**：封面/目录/章节页克隆自模板，母版自带底部蓝条与校徽

完整设计决策与"为什么"见 [`references/design_rules.md`](references/design_rules.md)。

## 作为 Agent 技能使用

把本目录放入 AI 编程助手的技能目录（如 Kimi Code 的 skills 路径），助手会在"做汇报 PPT / 工作汇报 / 武汉理工"等请求下自动触发，按 `SKILL.md` 的 6 步工作流执行：

```
需求确认 → 拆解源文档(≥90%覆盖) → 写大纲 JSON → 生成 → 机器自检 → PNG 逐页视觉自检
```
