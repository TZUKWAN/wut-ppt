# -*- coding: utf-8 -*-
"""Hard density checks for WUT body pages.

SVG mode (project path): checks information-unit counts.
Preview mode (directory with slide_XX.png): checks body-page spread and large
connected blank regions in the fixed-template body area.
"""
from pathlib import Path
import argparse
import json
import sys
import xml.etree.ElementTree as ET

from PIL import Image

CONTENT = (40, 76, 1240, 688)
GRID_COLS, GRID_ROWS = 12, 6
GRAPHIC_TAGS = {"rect", "circle", "ellipse", "polygon", "polyline", "path", "line", "image"}


def local_tag(element):
    return element.tag.rsplit("}", 1)[-1]


def audit_svg(project: Path, min_text: int, min_shapes: int):
    errors = []
    for svg in sorted((project / "svg_output").glob("*.svg")):
        root = ET.parse(svg).getroot()
        texts = [e for e in root.iter() if local_tag(e) == "text" and "".join(e.itertext()).strip()]
        shapes = [e for e in root.iter() if local_tag(e) in GRAPHIC_TAGS]
        if len(texts) < min_text:
            errors.append(f"{svg.name}: 文本信息单元 {len(texts)} < {min_text}，页面信息密度不足")
        if len(shapes) < min_shapes:
            errors.append(f"{svg.name}: 可见图形元素 {len(shapes)} < {min_shapes}，页面过于空疏")
    return errors


def body_slide_numbers(deck_path: Path | None, slide_count: int):
    if not deck_path or not deck_path.is_file():
        return set(range(1, slide_count + 1))
    config = json.loads(deck_path.read_text(encoding="utf-8"))
    number = 2  # fixed cover + TOC
    body = set()
    for chapter in config.get("chapters", []):
        number += 1  # fixed section page
        for _ in chapter.get("pages", []):
            number += 1
            body.add(number)
    return body


def blank_run(mask, cols, rows):
    seen = set()
    largest = 0
    for y in range(rows):
        for x in range(cols):
            if (x, y) in seen or mask[y][x]:
                continue
            stack, region = [(x, y)], 0
            seen.add((x, y))
            while stack:
                cx, cy = stack.pop()
                region += 1
                for nx, ny in ((cx - 1, cy), (cx + 1, cy), (cx, cy - 1), (cx, cy + 1)):
                    if 0 <= nx < cols and 0 <= ny < rows and (nx, ny) not in seen and not mask[ny][nx]:
                        seen.add((nx, ny))
                        stack.append((nx, ny))
            largest = max(largest, region)
    return largest


def audit_png(path: Path, max_blank_cells: int):
    with Image.open(path) as source:
        image = source.convert("RGB").crop(CONTENT)
    width, height = image.size
    pixels = image.load()
    nonwhite = [[False] * width for _ in range(height)]
    min_x, min_y, max_x, max_y = width, height, -1, -1
    ink = 0
    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y]
            if r < 240 or g < 240 or b < 240:
                nonwhite[y][x] = True
                ink += 1
                min_x, min_y = min(min_x, x), min(min_y, y)
                max_x, max_y = max(max_x, x), max(max_y, y)
    if max_x < 0:
        return ["整页正文为空白"], {"coverage_x": 0, "coverage_y": 0, "largest_blank": GRID_COLS * GRID_ROWS, "ink": 0}

    cell_occupied = []
    for row in range(GRID_ROWS):
        occupied_row = []
        y0, y1 = row * height // GRID_ROWS, (row + 1) * height // GRID_ROWS
        for col in range(GRID_COLS):
            x0, x1 = col * width // GRID_COLS, (col + 1) * width // GRID_COLS
            total = max(1, (x1 - x0) * (y1 - y0))
            count = sum(nonwhite[y][x] for y in range(y0, y1) for x in range(x0, x1))
            occupied_row.append(count / total >= 0.01)
        cell_occupied.append(occupied_row)
    largest = blank_run(cell_occupied, GRID_COLS, GRID_ROWS)
    metrics = {
        "coverage_x": (max_x - min_x + 1) / width,
        "coverage_y": (max_y - min_y + 1) / height,
        "largest_blank": largest,
        "ink": ink / (width * height),
    }
    errors = []
    if metrics["coverage_x"] < 0.90 or metrics["coverage_y"] < 0.85:
        errors.append(
            f"正文铺展不足：横向 {metrics['coverage_x']:.0%}、纵向 {metrics['coverage_y']:.0%}，"
            "内容未占满固定正文区"
        )
    if largest > max_blank_cells:
        errors.append(
            f"发现连续空白区域约 {largest}/{GRID_COLS * GRID_ROWS} 个网格，"
            f"超过上限 {max_blank_cells}；不得留下整片空白"
        )
    return errors, metrics


def main() -> None:
    parser = argparse.ArgumentParser(description="wut-ppt body-page density and blank-region audit")
    parser.add_argument("path", help="project directory or preview directory")
    parser.add_argument("--deck", help="deck.json used to skip fixed cover/TOC/section pages")
    parser.add_argument("--min-text", type=int, default=10)
    parser.add_argument("--min-shapes", type=int, default=18)
    parser.add_argument("--max-blank-cells", type=int, default=6)
    args = parser.parse_args()
    path = Path(args.path)
    errors = []
    if (path / "svg_output").is_dir():
        errors = audit_svg(path, args.min_text, args.min_shapes)
        for error in errors:
            print(f"[ERROR] {error}")
        if not errors:
            print(f"[OK] SVG 信息密度检查通过：每页至少 {args.min_text} 个文本单元、{args.min_shapes} 个图形元素")
    else:
        files = sorted(path.glob("slide_*.png"))
        if not files:
            raise SystemExit(f"未找到项目或 slide_XX.png: {path}")
        body = body_slide_numbers(Path(args.deck) if args.deck else None, len(files))
        for index, png in enumerate(files, 1):
            if index not in body:
                continue
            page_errors, metrics = audit_png(png, args.max_blank_cells)
            print(
                f"[DENSITY] slide_{index:02d}: 横向 {metrics['coverage_x']:.0%}，"
                f"纵向 {metrics['coverage_y']:.0%}，最大连续空白 {metrics['largest_blank']} 格"
            )
            errors.extend(f"{png.name}: {error}" for error in page_errors)
        for error in errors:
            print(f"[ERROR] {error}")
        if not errors:
            print("[OK] PowerPoint 正文铺展与空白区域检查通过")
    if errors:
        print(f"[FAIL] 信息密度检查未通过（{len(errors)} 项）")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
