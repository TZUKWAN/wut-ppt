# -*- coding: utf-8 -*-
"""
check_deck.py — wut-ppt 产出物自检（生成后必跑）

四道关：
  1. 占位符/模板提示语残留（页面标题、输入一章节标题、{{KEY}} 等）
  2. 字号下限：任何带文字 run 不得小于 12pt
  3. 色值合规：srgbClr 必须在设计调色板 + 模板自带色之内
  4. 疑似空页：整页形状数过少给出警告

用法: python scripts/check_deck.py 输出.pptx [--strict]
退出码: 0 通过, 1 不通过, 2 用法错误
"""
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(BASE, "assets", "template.pptx")

# 设计调色板（与 build_deck.py 设计常量一致）
PALETTE = {
    "00469A", "003371", "3D6FB0",  # 主蓝/深蓝/浅蓝(连接线)
    "FBC540", "9A7300",            # 金黄/深金
    "C00000",                      # 强调红
    "262626", "595959",            # 正文/辅助灰
    "FFFFFF", "F4F7FB", "EAF1F9",  # 白/卡片底/标签底
    "D9D9D9", "DCE6F5",            # 细线/横幅副行
}

DIRTY_MARKERS = [
    "此处输入一级标题", "演讲人姓名", "页面标题", "单击此处",
    "输入一章节标题", "输入二章节标题", "输入三章节标题",
    "输入四章节标题", "输入五章节标题", "{{", "}}",
]

MIN_PT = 12


def _template_colors():
    if not os.path.exists(TEMPLATE):
        return set()
    try:
        from pptx import Presentation
        from lxml import etree
        prs = Presentation(TEMPLATE)
        colors = set()
        for part in [prs.slide_masters[0].element] + [s._element for s in prs.slides]:
            xml = etree.tostring(part, encoding="unicode")
            colors.update(c.upper() for c in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', xml))
        return colors
    except Exception:
        return set()


def _walk(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)
        else:
            yield sh


def check(path):
    from pptx import Presentation
    from lxml import etree
    issues, warnings = [], []
    whitelist = PALETTE | _template_colors()
    prs = Presentation(path)

    for i, slide in enumerate(prs.slides):
        shapes = list(_walk(slide.shapes))
        if len(shapes) < 2:
            warnings.append(f"page[{i + 1}] 形状过少({len(shapes)})，疑似空页")

        for sh in shapes:
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text
            for m in DIRTY_MARKERS:
                if m in txt:
                    issues.append(f"page[{i + 1}] 占位符/模板提示语残留: {m!r}")
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.text.strip() and r.font.size is not None and r.font.size.pt < MIN_PT - 0.01:
                        issues.append(
                            f"page[{i + 1}] 字号 {r.font.size.pt:.1f}pt < {MIN_PT}pt: {r.text.strip()[:12]!r}")

        xml = etree.tostring(slide._element, encoding="unicode")
        for c in set(re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', xml)):
            if c.upper() not in whitelist:
                warnings.append(f"page[{i + 1}] 非调色板色值: #{c.upper()}")

    return issues, warnings


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    strict = "--strict" in sys.argv
    if not args or not os.path.exists(args[0]):
        print("用法: python scripts/check_deck.py 输出.pptx [--strict]")
        sys.exit(2)
    issues, warnings = check(args[0])
    print(f"=== wut-ppt 自检报告: {os.path.basename(args[0])} ===")
    for w in warnings[:10]:
        print(f"  ⚠️  {w}")
    if issues:
        for it in issues[:15]:
            print(f"  ❌ {it}")
        print(f"\n结论: 不通过（{len(issues)} 项违规，需返工）")
        sys.exit(1)
    if strict and warnings:
        print(f"\n结论: 严格模式不通过（{len(warnings)} 个警告）")
        sys.exit(1)
    print(f"✅ 通过（{len(warnings)} 个警告）")
    sys.exit(0)


if __name__ == "__main__":
    main()
