# -*- coding: utf-8 -*-
"""Audit deck-level layout diversity from design_spec.md and authored SVGs.

Hard rules:
- Layout ID, Geometry signature and Macro composition are required per slide.
- Each value must be unique within one deck.
- No two SVGs may share the same complete top-level module geometry.

The machine check is intentionally conservative. Synonymous compositions still
require the mandatory contact-sheet review by the main agent.
"""
from pathlib import Path
from collections import Counter
import argparse
import re
import sys
import xml.etree.ElementTree as ET

REQUIRED = (
    "Layout ID",
    "Geometry signature",
    "Macro composition",
    "Reading direction",
    "Visual center",
    "Rendering language",
    "Distinction",
)
UNIQUE = ("Layout ID", "Geometry signature", "Macro composition")


def norm(value: str) -> str:
    return re.sub(r"[\s`'\"“”‘’，,。；;：:]+", "", value).lower()


def parse_spec(path: Path):
    text = path.read_text(encoding="utf-8")
    matches = list(re.finditer(r"(?m)^#### Slide\s+(\d+)\s*-\s*(.+?)\s*$", text))
    pages = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        block = text[match.end():end]
        fields = {
            key.strip(): value.strip()
            for key, value in re.findall(r"(?m)^-\s+\*\*([^*]+)\*\*:\s*(.*?)\s*$", block)
        }
        pages.append({"number": match.group(1), "title": match.group(2), "fields": fields})
    return pages


def svg_fingerprint(path: Path):
    root = ET.parse(path).getroot()
    geometry = []
    full = []
    for group in [child for child in root if child.tag.rsplit("}", 1)[-1] == "g"]:
        bounds = group.get("data-pptx-bounds") or group.get("data-pptx-frame") or "unbounded"
        tags = Counter(element.tag.rsplit("}", 1)[-1] for element in group.iter())
        geometry.append(bounds)
        full.append((bounds, tuple(sorted(tags.items()))))
    return tuple(geometry), tuple(full)


def arrow_usage(path: Path):
    root = ET.parse(path).getroot()
    markers = sum(1 for element in root.iter() if element.get("marker-end") or element.get("marker-start"))
    presets = sum(
        1 for element in root.iter()
        if "Arrow" in (element.get("data-pptx-prst") or "") or (element.get("data-pptx-prst") or "") == "chevron"
    )
    return markers + presets


def audit(project: Path):
    errors, warnings = [], []
    pages = parse_spec(project / "design_spec.md")
    if not pages:
        errors.append("design_spec.md 未找到 #### Slide NN 页面块")
    seen = {field: {} for field in UNIQUE}
    rows = []
    for page in pages:
        fields = page["fields"]
        for field in REQUIRED:
            if not fields.get(field):
                errors.append(f"Slide {page['number']} 缺少字段: {field}")
        for field in UNIQUE:
            value = norm(fields.get(field, ""))
            if not value:
                continue
            if value in seen[field]:
                errors.append(
                    f"Slide {page['number']} 与 Slide {seen[field][value]} 的 {field} 重复: {fields[field]}"
                )
            else:
                seen[field][value] = page["number"]
        rows.append((page["number"], page["title"], fields))

    svgs = sorted((project / "svg_output").glob("*.svg"))
    if pages and len(svgs) != len(pages):
        errors.append(f"页面规划数({len(pages)})与 SVG 数({len(svgs)})不一致")
    fingerprints = []
    for index, svg in enumerate(svgs):
        try:
            fingerprints.append((svg.name, *svg_fingerprint(svg)))
            if index < len(pages):
                fields = pages[index]["fields"]
                arrows = arrow_usage(svg)
                semantics = fields.get("Arrow semantics", "")
                if arrows and not semantics:
                    errors.append(
                        f"{svg.name}: 使用 {arrows} 处箭头但未填写 Arrow semantics；"
                        "必须说明来源→目标及业务语义，禁止装饰性箭头"
                    )
                elif arrows and "→" not in semantics and "->" not in semantics:
                    errors.append(f"{svg.name}: Arrow semantics 必须包含明确的 来源→目标 关系")
                if arrows and fields.get("Macro composition") == "table-ledger":
                    if not any(word in semantics for word in ("交接", "映射", "转化", "流转", "支撑")):
                        errors.append(f"{svg.name}: 表格中的箭头缺少交接/映射/转化等真实业务语义")
        except ET.ParseError as exc:
            errors.append(f"{svg.name}: XML 解析失败: {exc}")
    for i, (name_a, geometry_a, full_a) in enumerate(fingerprints):
        for name_b, geometry_b, full_b in fingerprints[i + 1:]:
            if geometry_a and geometry_a == geometry_b:
                errors.append(f"{name_a} 与 {name_b} 的顶层模块几何完全重复")
            elif full_a and full_a == full_b:
                errors.append(f"{name_a} 与 {name_b} 的顶层模块结构完全重复")

    validation = project / "validation"
    validation.mkdir(exist_ok=True)
    lines = [
        "# Deck Visual Map",
        "",
        "> Machine uniqueness audit is necessary but not sufficient; review the contact sheet for synonymous compositions.",
        "",
        "| Page | Layout ID | Geometry signature | Macro composition | Reading direction | Visual center | Rendering | Distinction |",
        "|---|---|---|---|---|---|---|---|",
    ]
    for number, title, fields in rows:
        lines.append("| P{} {} | {} | {} | {} | {} | {} | {} | {} |".format(
            number, title, *(fields.get(field, "—").replace("|", "\\|") for field in REQUIRED)
        ))
    (validation / "deck_visual_map.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    return pages, errors, warnings


def main() -> None:
    parser = argparse.ArgumentParser(description="wut-ppt deck layout uniqueness audit")
    parser.add_argument("project")
    args = parser.parse_args()
    pages, errors, warnings = audit(Path(args.project))
    for warning in warnings:
        print(f"[WARN] {warning}")
    for error in errors:
        print(f"[ERROR] {error}")
    if errors:
        print(f"[FAIL] 版式唯一性检查未通过（{len(errors)} 项）")
        raise SystemExit(1)
    print(f"[OK] 版式唯一性检查通过：{len(pages)} 页，Layout ID / 几何签名 / 宏观构图均唯一")
    print("[NEXT] 仍须生成 contact_sheet.png，人工排除近义同构版式")


if __name__ == "__main__":
    main()
