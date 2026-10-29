# -*- coding: utf-8 -*-
"""Clone the fixed WUT template and inject native SVG-converted body content.

Usage:
  python scripts/assemble_deck.py deck.json svg_native.pptx output.pptx

deck.json:
{
  "title": "封面主标题", "speaker": "汇报人",
  "chapters": [{"name": "章节名", "pages": [{"title": "正文页标题"}, ...]}]
}
The source PPTX slides must contain only body content, in the same order as
`chapters[].pages[]`. Cover/TOC/section/title/master never come from SVG.
"""
from copy import deepcopy
from io import BytesIO
from pathlib import Path
import json
import sys
import zipfile

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

BASE = Path(__file__).resolve().parent.parent
TEMPLATE = BASE / "assets" / "template.pptx"
CH_NUM = "一二三四五"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CONTENT = {"left": 40, "top": 76, "right": 1240, "bottom": 688}


def replace_deep(slide, mapping):
    def process(shape):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in shape.shapes:
                process(sub)
            return
        if not shape.has_text_frame:
            return
        for paragraph in shape.text_frame.paragraphs:
            runs = list(paragraph.runs)
            if not runs:
                continue
            full = "".join(run.text for run in runs)
            new = full
            for old, value in mapping.items():
                new = new.replace(old, str(value))
            if new != full:
                runs[0].text = new
                for run in runs[1:]:
                    run.text = ""
    for shape in slide.shapes:
        process(shape)


def clone(prs, source):
    slide = prs.slides.add_slide(source.slide_layout)
    for placeholder in list(slide.placeholders):
        slide.shapes._spTree.remove(placeholder._element)
    for shape in source.shapes:
        slide.shapes._spTree.append(deepcopy(shape._element))
    return slide


def copy_relationships(element, source_slide, target_slide):
    for node in element.iter():
        for attr, old_rid in list(node.attrib.items()):
            if not attr.startswith(REL_NS) or not old_rid.startswith("rId"):
                continue
            rel = source_slide.part.rels[old_rid]
            if rel.reltype == RT.IMAGE:
                _, new_rid = target_slide.part.get_or_add_image_part(BytesIO(rel.target_part.blob))
            elif rel.is_external:
                new_rid = target_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                raise ValueError(f"不支持复制的内部关系: {rel.reltype}")
            node.set(attr, new_rid)


def theme_colors(native_pptx):
    """Read the source deck theme so copied shapes keep their exact RGB colors."""
    from lxml import etree
    with zipfile.ZipFile(native_pptx) as package:
        root = etree.fromstring(package.read("ppt/theme/theme1.xml"))
    colors = {}
    for role in root.findall(".//a:clrScheme/*", {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}):
        if len(role):
            value = role[0].get("val") or role[0].get("lastClr")
            if value:
                colors[role.tag.rsplit("}", 1)[-1]] = value.upper()
    return colors


def freeze_theme_colors(element, colors):
    """Prevent the destination template theme from recoloring imported content."""
    for node in element.iter():
        if node.tag == qn("a:schemeClr") and node.get("val") in colors:
            node.tag = qn("a:srgbClr")
            node.set("val", colors[node.get("val")])
        elif node.tag == qn("a:sysClr") and node.get("lastClr"):
            node.tag = qn("a:srgbClr")
            node.set("val", node.get("lastClr"))
            node.attrib.pop("lastClr", None)


def inject_content(target_slide, source_slide, colors):
    for shape in source_slide.shapes:
        left = shape.left / 914400 * 96
        top = shape.top / 914400 * 96
        right = (shape.left + shape.width) / 914400 * 96
        bottom = (shape.top + shape.height) / 914400 * 96
        if left < CONTENT["left"] - 2 or top < CONTENT["top"] - 2 or right > CONTENT["right"] + 2 or bottom > CONTENT["bottom"] + 2:
            raise ValueError(
                f"正文内容越过固定模板安全区: {shape.name} "
                f"({left:.1f},{top:.1f},{right:.1f},{bottom:.1f})"
            )
        element = deepcopy(shape._element)
        freeze_theme_colors(element, colors)
        copy_relationships(element, source_slide, target_slide)
        target_slide.shapes._spTree.append(element)


def build(config, native_pptx, output):
    source = Presentation(native_pptx)
    colors = theme_colors(native_pptx)
    expected = [page for chapter in config["chapters"] for page in chapter["pages"]]
    if len(source.slides) != len(expected):
        raise ValueError(f"原生 SVG 页数({len(source.slides)})与正文规划页数({len(expected)})不一致")

    prs = Presentation(TEMPLATE)
    templates = list(prs.slides)
    cover = clone(prs, templates[0])
    replace_deep(cover, {"此处输入一级标题": config["title"], "演讲人姓名": config["speaker"]})

    chapters = config["chapters"]
    toc = clone(prs, templates[1])
    replace_deep(toc, {f"输入{CH_NUM[i]}章节标题": chapter["name"] for i, chapter in enumerate(chapters[:5])})
    if len(chapters) < 5:
        extras = []
        for i in range(len(chapters), 5):
            extras += [f"输入{CH_NUM[i]}章节标题", CH_NUM[i]]
        for shape in list(toc.shapes):
            if shape.has_text_frame and shape.text_frame.text.strip() in extras:
                toc.shapes._spTree.remove(shape._element)

    source_index = 0
    for chapter_index, chapter in enumerate(chapters):
        divider = clone(prs, templates[2])
        replace_deep(divider, {"输入一章节标题": f"{CH_NUM[chapter_index]}、{chapter['name']}"})
        for page in chapter["pages"]:
            slide = clone(prs, templates[3])
            replace_deep(slide, {"页面标题": page["title"]})
            inject_content(slide, source.slides[source_index], colors)
            source_index += 1
            print(f"  OK {page['title']}")

    for slide_id in list(prs.slides._sldIdLst)[:5]:
        prs.slides._sldIdLst.remove(slide_id)
        prs.part.drop_rel(slide_id.rId)
    prs.save(output)
    print(f"OK {output}（{len(prs.slides)}页，固定模板＋原生正文）")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit("用法: python scripts/assemble_deck.py deck.json svg_native.pptx output.pptx")
    build(json.loads(Path(sys.argv[1]).read_text(encoding="utf-8")), sys.argv[2], sys.argv[3])
