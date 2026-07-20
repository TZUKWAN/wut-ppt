---
name: wut-ppt
description: 武汉理工大学行政汇报 PPT 生成技能。以张林定制模板为唯一载体（校徽蓝#00469A+金#FBC540+红#C00000），按"大纲 JSON + 12 种 zone 版式"生成高密度逻辑图风 PPT：每个要点独立形状包裹、箭头/时间轴逻辑引导、全文≥12pt、自动填满无大片空白。触发词：做汇报PPT、工作汇报、进展汇报、XX会汇报、阶段总结、行动计划汇报、武汉理工、智链。
---

# 武汉理工大学行政汇报 PPT 技能 (wut-ppt)

质量标准见 `references/design_rules.md`（含失败教训分析）。
**每一页都要达到"高密度逻辑图"水准，不允许退回关键词漂浮或大色块。**

## 硬规则（不可违反）

1. **载体唯一**：封面/目录/章节页/正文标题一律克隆 `assets/template.pptx`，母版自带底部蓝条+校徽
2. **颜色三件套**：蓝 `#00469A` 主色、金 `#FBC540` 点缀、红 `#C00000` 仅结论/关键数字；其余只用黑白灰（详见 `references/design_rules.md`）
3. **字号下限 12pt**：任何带文字的 run 不得小于 12pt（check_deck 强制检查）
4. **每点独立形状**：要点必须包进独立卡片 + 序号圆徽，禁止纯文字列表
5. **逻辑引导**：递进用箭头链、时间用时间轴、并列用连接线、映射用箭头，形状用圆/圆角/六边形/平行四边形/CHEVRON
6. **高信息密度**：要点 = 引导语 + 完整解释句（30–65 字），每页 300–600 字，覆盖源文档 90%+
7. **零大片空白**：zone 垂直流自动填满，生成后逐页目检确认

## 工作流（6 步，缺一不可）

```
① 需求确认：汇报场合、受众、章节数、必须出现的数据点/专名
② 拆源文档：每段原文映射到页（覆盖率 ≥90%）
③ 写大纲 JSON：按 references/outline_contract.md 选 zone 配方
   （页面配方表直接套用）
④ 生成：python scripts/build_deck.py 大纲.json 输出.pptx
⑤ 机器自检：python scripts/check_deck.py 输出.pptx
⑥ 视觉自检：python scripts/export_preview.py 输出.pptx 预览目录
   → 逐页看 PNG，对照下方验收清单；不过则改 JSON 回到 ④
```

**第 ⑥ 步人眼看图不可省**——断词、拥挤、空白、引导缺失，机器查不出来。

## 逐页验收清单

- [ ] 封面/目录/章节页无占位符残留（"输入一章节标题"之类）
- [ ] 每页正文填满，无大面积空白
- [ ] 每个要点都在独立形状内，有序号/徽章
- [ ] 页面有逻辑引导形状（箭头/连接线/时间轴）
- [ ] 文字密度足够，无关键词漂浮
- [ ] 颜色未超出三件套+中性色

## 命令速查

```bash
python scripts/build_deck.py 大纲.json 输出.pptx   # 生成
python scripts/check_deck.py 输出.pptx             # 机器自检（占位符/字号/色值/空页）
python scripts/export_preview.py 输出.pptx out     # 导出每页 PNG（需本机装 PowerPoint）
```

## 文件清单

```
wut-ppt/
├── SKILL.md                        本文件
├── assets/template.pptx            张林定制模板（唯一载体，5 样板页+母版）
├── scripts/
│   ├── build_deck.py               生成器（zone 垂直流布局引擎，12 种版式）
│   ├── check_deck.py               机器自检
│   └── export_preview.py           PNG 导出（视觉自检）
├── references/
│   ├── outline_contract.md         大纲 JSON 契约：12 种 zone 用法 + 页面配方
│   ├── design_rules.md             设计系统 + 失败教训分析（改设计前必读）
│   └── terminology.md              武理工术语词典（专名不可写错）
```

## 大纲 JSON 速览

```json
{"title": "...", "speaker": "...",
 "chapters": [{"name": "章节名",
   "pages": [{"title": "页面标题",
     "zones": [
       {"type": "lead", "text": "导语段……"},
       {"type": "cards", "cards": [{"head": "维度一", "items": [
           {"lead": "引导语", "text": "完整解释句……"}]}]},
       {"type": "note", "text": "结论条……"}]}]}],
 "closing": {"line1": "以上汇报，敬请批评指正", "line2": "汇报人：XXX"}}
```

12 种 zone：`lead / cards / flow_chain / timeline / compare_rows / badge_grid / shape_row / chips / kv_rows / banner / table / note`。
标记语法：`**蓝粗强调**`、`[[红粗关键数字]]`。
完整字段与选型指南 → `references/outline_contract.md`。
