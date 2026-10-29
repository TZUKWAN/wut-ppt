# 固定结构配置、正文视觉地图与 SVG 契约

## 一、项目结构

```text
design_spec.md       正文内容、逐页设计和全稿正文视觉地图
spec_lock.md         颜色、字体、页面节奏和结构锁
deck.json            固定封面、汇报人、章节和正文页标题
svg_output/          AI 逐页手写的正文 SVG（唯一正文视觉源）
notes/total.md       可选演讲备注
images/              模板及项目图片资源
validation/          技术报告和 deck_visual_map.md
```

## 二、固定结构配置 deck.json

```json
{
  "title": "封面主标题",
  "speaker": "汇报人",
  "chapters": [
    {"name": "第一章名称", "pages": [
      {"title": "正文页标题"},
      {"title": "下一页标题"}
    ]}
  ]
}
```

- 最多五章，固定封面、目录和章节页由模板自动生成。
- 所有章节的 `pages` 顺序必须与 `svg_output/NN_*.svg` 完全一致。
- 不得在 `deck.json` 中配置自定义封面、底版或结构页版式。

## 三、每个正文页必填字段

```markdown
#### Slide 01 - 页面短名
- **Audience move**: 受众看完应理解或接受什么
- **Composition**: 正文主构图、阅读顺序、图文比例
- **Title**: 正文页标题（与 deck.json 一致）
- **Core message**: 本页唯一核心判断
- **Relationships**: 层级/顺序/汇聚/交叠/反馈/支撑等真实关系
- **Layout ID**: L01（全稿唯一）
- **Geometry signature**: 页面任务 → 宏观构图 → 阅读方向 → 视觉中心 → 渲染语言（全稿唯一）
- **Macro composition**: composition_library.md 中的 ID 或 custom-具体名称（全稿唯一）
- **Reading direction**: 例如自上而下 / 中心向四周
- **Visual center**: 页面第一视觉落点
- **Rendering language**: admin-editorial / data-journalism / engineering-blueprint 等
- **Distinction**: 与此前所有正文页在宏观构图上的本质区别
- **Arrow semantics**: 使用箭头时必填“来源→目标＋交接/映射/转化/支撑语义”
- **Content**: 呈现在正文区的最终内容，而非写作提示
```

- 换内容、颜色、图标、节点数不能作为 `Distinction`。
- 任何箭头，包括原生箭头预设，都必须填写 `Arrow semantics`；没有箭头可不填。
- 表格、泳道没有真实交接或映射时，不得使用箭头。
- 页面节奏只用 dense，不得设置 breathing 正文页。

## 四、正文 SVG 基本要求

```xml
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720"
     data-pptx-page-role="content"
     font-family="Microsoft YaHei, Arial, sans-serif">
  <g id="main-diagram" data-pptx-bounds="40 76 1200 612">
    <!-- 在整个正文安全区展开主图、关系线、节点和说明 -->
  </g>
</svg>
```

- 正文安全区：`x=40–1240，y=76–688`。
- 不得包含封面、目录、章节页、页标题、底部蓝条、校徽或整页品牌背景。
- 每个逻辑模块有唯一语义 `id` 和正数 `data-pptx-bounds`。
- 任何显式 `font-size` 不小于 16px。
- 每页至少 10 个文本信息单元、18 个可见图形元素；不得以无意义形状凑数。
- 禁止 `<style>`、`class`、`foreignObject`、`textPath`、脚本、动画和外部 CSS。

## 五、全页填满与箭头语义

- 内容必须铺展到正文区四边，禁止留下连续空白区域。
- 通过真实子要点、解释、任务、指标、交付物和输入输出补足，不得只拉大容器。
- 箭头只能用于真实递进、映射、交接、转化、支撑或反馈。
- 表格以网格和单元格内容为主；没有业务交接时不画箭头。
- 中心圆、环形图和泳道不能成为信息稀薄的默认答案。

## 六、换行与可编辑性

同一逻辑段落使用一个 `<text>`，后续行使用相同 `x` 和正数 `dy`：

```xml
<text x="730" y="100" font-size="20" fill="#262626">第一行完整表述，
  <tspan x="730" dy="30">第二行继续表述，</tspan>
  <tspan x="730" dy="30">第三行结束表述。</tspan>
</text>
```

- 行距通常 28–32px（20px 正文）。
- 非定位 `<tspan>` 可用于同一句内强调。
- 窄形状不要硬塞长词；需要竖排时逐字分行。
- 导出器不会替 AI 自动换行或缩小文字。

## 七、完成定义

1. `wut_svg.py check` 通过：规划字段、唯一性、箭头语义、SVG 几何、字号和静态信息密度。
2. `wut_svg.py export` 完成原生转换、固定模板组装和 PPTX 自检。
3. `wut_svg.py verify` 通过：PNG 正文横向铺展 ≥90%、纵向 ≥85%，12×6 网格最大连续空白 ≤6 格。
4. 主 Agent 打开 `contact_sheet.png` 和关键单页，排除近义同构、局部空白、无意义箭头、遮挡和断词。
